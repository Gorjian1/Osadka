using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace Osadka.ExcelDetection
{
    public sealed class ExcelStructureDetector
    {
        public sealed record DetectedCycle(int Index, int ColumnStart, string Title);
        public sealed record RowRange(int StartRow, int EndRow);
        public sealed record DetectedObject(
            int ObjectIndex,
            string Name,
            int IdColumn,
            int HeaderRow,
            int SubHeaderRow,
            IReadOnlyList<RowRange> Blocks,
            IReadOnlyList<DetectedCycle> Cycles);

        public static List<DetectedObject> Detect(IXLWorksheet ws)
        {
            var used = ws.RangeUsed();
            if (used == null) return new List<DetectedObject>();

            int r1 = used.FirstRow().RowNumber();
            int r2 = used.LastRow().RowNumber();
            int c1 = used.FirstColumn().ColumnNumber();
            int c2 = used.LastColumn().ColumnNumber();

            var candidates = new List<(int HeaderRow, int IdCol, int SubHeaderRow, string ContextKey)>();

            // 1) ищем все кандидаты заголовка ID (склейка до 3 строк вертикально)
            for (int r = r1; r <= r2; r++)
            {
                for (int c = c1; c <= c2; c++)
                {
                    string s0 = T(ws, r, c, r1, r2, c1, c2);
                    string s01 = (s0 + " " + T(ws, r + 1, c, r1, r2, c1, c2)).Trim();
                    string s012 = (s01 + " " + T(ws, r + 2, c, r1, r2, c1, c2)).Trim();
                    string s_10 = (T(ws, r - 1, c, r1, r2, c1, c2) + " " + s0).Trim();

                    bool hit = IsIdHeader(s0) || IsIdHeader(s01) || IsIdHeader(s012) || IsIdHeader(s_10);
                    if (!hit) continue;

                    int headerRow = IsIdHeader(s_10) ? Math.Max(r - 1, r1) : r;
                    var subHeader = FindSubHeaderRow(ws, headerRow, c, r1, r2, c1, c2);
                    if (subHeader == null) continue;

                    string context = BuildContextSignature(ws, headerRow, c, r1, r2, c1, c2);
                    if (!HasIdsBelow(ws, subHeader.Value + 1, c, r1, r2, c1, c2)) continue;

                    candidates.Add((headerRow, c, subHeader.Value, context));
                }
            }
            if (candidates.Count == 0) return new List<DetectedObject>();

            // 2) дедуп по строкам: оставляем самый левый ID-столбец на каждой строке шапки
            var dedup = candidates
                .GroupBy(x => x.HeaderRow)
                .Select(g => g.OrderBy(x => x.IdCol).First())
                .OrderBy(x => x.HeaderRow)
                .ToList();

            // 3) группируем кандидатов по «контекстной сигнатуре»
            var grouped = dedup
                .GroupBy(x => x.ContextKey)
                .OrderBy(g => g.Min(v => v.HeaderRow))
                .ToList();

            var objects = new List<DetectedObject>();
            int objIndex = 0;

            foreach (var g in grouped)
            {
                var first = g.OrderBy(x => x.HeaderRow).First();

                var cycles = DetectCycles(ws, first.HeaderRow, first.SubHeaderRow, first.IdCol, r1, r2, c1, c2);
                if (cycles.Count == 0) continue;

                // склеиваем блоки строк данных
                var blocks = new List<RowRange>();
                foreach (var cand in g.OrderBy(x => x.HeaderRow))
                {
                    var rr = DetectDataRange(ws, cand.SubHeaderRow, cand.IdCol, r1, r2, c1, c2);
                    if (rr.EndRow < rr.StartRow) continue;

                    if (blocks.Count == 0)
                        blocks.Add(rr);
                    else
                    {
                        var last = blocks[^1];
                        if (rr.StartRow <= last.EndRow + 3)
                            blocks[^1] = new RowRange(last.StartRow, Math.Max(last.EndRow, rr.EndRow));
                        else
                            blocks.Add(rr);
                    }
                }
                if (blocks.Count == 0) continue;

                objIndex++;
                string name = MakeObjectName(ws, first.HeaderRow, first.IdCol, first.SubHeaderRow, r1, r2, c1, c2, objIndex);
                objects.Add(new DetectedObject(objIndex, name, first.IdCol, first.HeaderRow, first.SubHeaderRow, blocks, cycles));
            }

            return objects;
        }

        // ---------- helpers ----------

        private static string T(IXLWorksheet ws, int r, int c, int r1, int r2, int c1, int c2)
            => (r < r1 || r > r2 || c < c1 || c > c2) ? string.Empty : ws.Cell(r, c).GetString()?.Trim() ?? string.Empty;

        private static bool IsIdHeader(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            var t = Regex.Replace(s.ToLowerInvariant(), @"\s+", " ");
            return Regex.IsMatch(t, @"(^|[\s:])№\s*(мар(ка|ки)?|точ(ка|ки)?)\b")
                || Regex.IsMatch(t, @"\bномер\s+(мар(ки|ка)?|точ(ки|ка)?)\b");
        }

        private static bool LooksLikeCycleLabel(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            var t = Regex.Replace(s.ToLowerInvariant(), @"\s+", " ");
            return t.Contains("отметк") || t.Contains("осадк") || t.Contains("общая осад");
        }

        private static int? FindSubHeaderRow(IXLWorksheet ws, int headerRow, int idCol, int r1, int r2, int c1, int c2)
        {
            for (int r = headerRow; r <= Math.Min(headerRow + 10, r2); r++)
            {
                int hits = 0;
                for (int c = idCol + 1; c <= Math.Min(idCol + 60, c2); c++)
                    if (LooksLikeCycleLabel(T(ws, r, c, r1, r2, c1, c2))) hits++;

                if (hits >= 3) return r;
            }
            return null;
        }

        private static bool HasIdsBelow(IXLWorksheet ws, int startRow, int idCol, int r1, int r2, int c1, int c2)
        {
            int cnt = 0, emptyStreak = 0;
            for (int r = startRow; r <= Math.Min(startRow + 300, r2); r++)
            {
                string s = T(ws, r, idCol, r1, r2, c1, c2);
                if (string.IsNullOrWhiteSpace(s))
                {
                    emptyStreak++;
                    if (emptyStreak >= 10 && cnt > 0) break;
                    continue;
                }
                emptyStreak = 0;
                var low = s.ToLowerInvariant();
                if (IsIdHeader(s) || low.Contains("расположенных") || low.Contains("ведомость"))
                {
                    if (cnt == 0) continue;
                    else break;
                }
                cnt++;
                if (cnt >= 3) return true;
            }
            return cnt >= 3;
        }

        private static string BuildContextSignature(IXLWorksheet ws, int headerRow, int idCol, int r1, int r2, int c1, int c2)
        {
            var lines = new List<string>();
            for (int r = Math.Max(r1, headerRow - 3); r <= headerRow - 1; r++)
            {
                var rowCells = new List<string>();
                for (int c = Math.Max(c1, idCol - 8); c <= Math.Min(idCol + 40, c2); c++)
                {
                    var t = T(ws, r, c, r1, r2, c1, c2);
                    if (!string.IsNullOrWhiteSpace(t)) rowCells.Add(t);
                }
                if (rowCells.Count > 0) lines.Add(string.Join(" ", rowCells));
            }
            var raw = string.Join(" | ", lines);
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            return Regex.Replace(raw.ToLowerInvariant(), @"\s+", " ").Trim();
        }

        private static List<DetectedCycle> DetectCycles(IXLWorksheet ws, int headerRow, int subHeaderRow, int idCol, int r1, int r2, int c1, int c2)
        {
            var result = new List<DetectedCycle>();
            int c = idCol + 1;
            int idx = 0;
            while (c <= c2 - 2)
            {
                var l1 = T(ws, subHeaderRow, c, r1, r2, c1, c2);
                var l2 = T(ws, subHeaderRow, c + 1, r1, r2, c1, c2);
                var l3 = T(ws, subHeaderRow, c + 2, r1, r2, c1, c2);

                int labels = 0;
                if (LooksLikeCycleLabel(l1)) labels++;
                if (LooksLikeCycleLabel(l2)) labels++;
                if (LooksLikeCycleLabel(l3)) labels++;

                if (labels >= 2)
                {
                    idx++;
                    string title = T(ws, headerRow, c, r1, r2, c1, c2);
                    if (string.IsNullOrWhiteSpace(title)) title = $"Цикл {idx}";
                    result.Add(new DetectedCycle(idx, c, title));
                    c += 3;
                }
                else
                    c++;
            }
            return result;
        }

        private static RowRange DetectDataRange(IXLWorksheet ws, int subHeaderRow, int idCol, int r1, int r2, int c1, int c2)
        {
            int start = subHeaderRow + 1;
            while (start <= r2 && string.IsNullOrWhiteSpace(T(ws, start, idCol, r1, r2, c1, c2)))
                start++;

            int end = start, emptyStreak = 0;
            for (int r = start; r <= r2; r++)
            {
                string id = T(ws, r, idCol, r1, r2, c1, c2);
                if (string.IsNullOrWhiteSpace(id))
                {
                    emptyStreak++;
                    if (emptyStreak >= 5) break;
                    continue;
                }
                emptyStreak = 0;
                var low = id.ToLowerInvariant();
                if (IsIdHeader(id) || low.Contains("расположенных") || low.Contains("ведомость"))
                    break;
                end = r;
            }
            if (end < start) end = start;
            return new RowRange(start, end);
        }

        private static string MakeObjectName(IXLWorksheet ws, int headerRow, int idCol, int subHeaderRow, int r1, int r2, int c1, int c2, int fallbackIndex)
        {
            for (int r = Math.Max(r1, headerRow - 3); r <= headerRow - 1; r++)
            {
                var rowText = string.Join(" ", Enumerable.Range(c1, c2 - c1 + 1)
                    .Select(cc => T(ws, r, cc, r1, r2, c1, c2))
                    .Where(t => !string.IsNullOrWhiteSpace(t)));

                var low = rowText.ToLowerInvariant();
                if (low.Contains("грунтов") || low.Contains("котлован") || low.Contains("инженерн") || low.Contains("окруж"))
                    return rowText.Trim();
                if (low.Contains("по адресу") || low.Contains("адрес"))
                    return rowText.Trim();
            }
            return $"{ws.Name} — Объект {fallbackIndex}";
        }
    }
}
