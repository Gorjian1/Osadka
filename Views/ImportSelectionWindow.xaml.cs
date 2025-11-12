// ReSharper disable InconsistentNaming
using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

namespace Osadka.Views
{
    public partial class ImportSelectionWindow : Window
    {
        public ImportSelectionWindow(XLWorkbook wb)
        {
            InitializeComponent();
            DataContext = new ImportSelectionVM(wb);
        }

        private void OkClick(object? sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        // === Публичные геттеры для RawDataViewModel ===
        public WorksheetItem? SelectedWorksheet => (DataContext as ImportSelectionVM)?.SelectedWorksheet;
        public List<(int Row, IXLCell Cell)> ObjectHeaders => (DataContext as ImportSelectionVM)?.ObjectHeaders ?? new();
        public List<int> CycleStarts => (DataContext as ImportSelectionVM)?.CycleStarts ?? new();
        public int SelectedObjectIndex => (DataContext as ImportSelectionVM)?.SelectedObjectIndex ?? 1; // 1-based
        public int SelectedCycleIndex => (DataContext as ImportSelectionVM)?.SelectedCycleIndex ?? 1; // 1-based
    }

    public class ImportSelectionVM : ObservableObject
    {
        private readonly XLWorkbook _wb;

        public List<WorksheetItem> Worksheets { get; } = new();

        private WorksheetItem? _selectedWorksheet;
        public WorksheetItem? SelectedWorksheet
        {
            get => _selectedWorksheet;
            set
            {
                if (SetProperty(ref _selectedWorksheet, value))
                {
                    LoadObjects();
                }
            }
        }

        public List<ObjectItem> Objects { get; private set; } = new();

        private ObjectItem? _selectedObject;
        public ObjectItem? SelectedObject
        {
            get => _selectedObject;
            set
            {
                if (SetProperty(ref _selectedObject, value))
                {
                    LoadCycles();
                }
            }
        }

        public List<CycleItem> Cycles { get; private set; } = new();

        private CycleItem? _selectedCycle;
        public CycleItem? SelectedCycle
        {
            get => _selectedCycle;
            set => SetProperty(ref _selectedCycle, value);
        }

        // Для передачи обратно во VM
        public List<(int Row, IXLCell Cell)> ObjectHeaders { get; private set; } = new();
        public List<int> CycleStarts { get; private set; } = new();

        // Выдаём индексы 1-based (как ожидает внутренняя нумерация слева-направо)
        public int SelectedObjectIndex => SelectedObject?.Index ?? 1;
        public int SelectedCycleIndex => SelectedCycle?.Index ?? 1;

        public ImportSelectionVM(XLWorkbook wb)
        {
            _wb = wb;

            var preferred = wb.Worksheets
                              .Select((s, i) => new WorksheetItem
                              {
                                  Index = i + 1,
                                  Name = s.Name,
                                  Sheet = s
                              })
                              .OrderByDescending(w => Regex.IsMatch(w.Name, "окруж", RegexOptions.IgnoreCase) ? 1 : 0)
                              .ThenBy(w => w.Index)
                              .ToList();

            Worksheets.AddRange(preferred);
            SelectedWorksheet = Worksheets.FirstOrDefault();
        }

        private void LoadObjects()
        {
            Objects = new();
            ObjectHeaders = new();
            SelectedObject = null;
            Cycles = new();
            SelectedCycle = null;
            CycleStarts = new();

            if (SelectedWorksheet == null) return;

            var ws = SelectedWorksheet.Sheet;

            ObjectHeaders = FindObjectHeadersV2(ws);

            var list = ObjectHeaders
                .Select((hdr, i) =>
                {
                    int leftCol = hdr.Cell.Address.ColumnNumber;
                    int rightCol = FindRightIdColumnInRow(ws, hdr.Row);
                    int firstDataRow = FindSubHeaderRow(ws, hdr.Row, leftCol) + 1;
                    int countRows = CountNumericRows(ws, firstDataRow, leftCol, rightCol);

                    return new ObjectItem
                    {
                        Index = i + 1,
                        HeaderRow = hdr.Row,
                        IdLeftColumn = leftCol,
                        IdRightColumn = rightCol,
                        RowsCountHint = countRows
                    };
                })
                .ToList();

            Objects = list;
            OnPropertyChanged(nameof(Objects));
            SelectedObject = Objects.FirstOrDefault();
        }

        private void LoadCycles()
        {
            Cycles = new();
            SelectedCycle = null;
            CycleStarts = new();

            if (SelectedWorksheet == null || SelectedObject == null) return;

            var ws = SelectedWorksheet.Sheet;
            int headerRow = SelectedObject.HeaderRow;
            int idColLeft = SelectedObject.IdLeftColumn;

            int subHeaderRow = FindSubHeaderRow(ws, headerRow, idColLeft);
            var starts = FindCycleStartsV2(ws, headerRow, subHeaderRow, idColLeft); // ASC по колонке

            // порядок = как в ComboBox (правый цикл первый)
            CycleStarts = starts
                .OrderByDescending(s => s.StartColumn)
                .Select(s => s.StartColumn)
                .ToList();

            // Внутренняя нумерация — слева-направо (ASC) для подписей
            var ordinalByCol = starts
                .Select((s, idx) => (s.StartColumn, Ordinal: idx + 1))
                .ToDictionary(x => x.StartColumn, x => x.Ordinal);

            // Отображение пользователю — справа-налево (часто самый правый — последний цикл)
            // Index должен соответствовать позиции в DESC-списке, чтобы совпадать с ReadAllObjects
            Cycles = starts
                .OrderByDescending(c => c.StartColumn)
                .Select((s, idx) => new CycleItem
                {
                    Index = idx + 1, // Позиция в DESC-списке (соответствует cycIdx в ReadAllObjects)
                    StartColumn = s.StartColumn,
                    Label = string.IsNullOrWhiteSpace(s.Label) ? $"Цикл {ordinalByCol[s.StartColumn]}" : s.Label
                })
                .ToList();

            OnPropertyChanged(nameof(Cycles));
            SelectedCycle = Cycles.FirstOrDefault();
        }

        // === Поиск заголовка «№ точки/№ марки» слева и справа ===
        private static List<(int Row, IXLCell Cell)> FindObjectHeadersV2(IXLWorksheet sheet)
        {
            var list = new List<(int Row, IXLCell Cell)>();
            var used = sheet.RangeUsed();
            if (used == null) return list;

            foreach (var row in used.Rows())
            {
                var candidates = row.Cells()
                    .Where(c =>
                    {
                        var s = c.GetString();
                        return Regex.IsMatch(s, @"^\s*№\s*(точки|мар\w*)", RegexOptions.IgnoreCase);
                    })
                    .ToList();

                if (candidates.Count >= 1)
                {
                    var leftMost = candidates.OrderBy(c => c.Address.ColumnNumber).First();
                    list.Add((row.RowNumber(), leftMost));
                }
            }
            return list;
        }

        private static int FindRightIdColumnInRow(IXLWorksheet sheet, int headerRow)
        {
            var row = sheet.Row(headerRow);
            var right = row.Cells()
                .Where(c => Regex.IsMatch(c.GetString(), @"^\s*№\s*(точки|мар\w*)", RegexOptions.IgnoreCase))
                .OrderByDescending(c => c.Address.ColumnNumber)
                .FirstOrDefault();
            return right?.Address.ColumnNumber ?? FindFallbackRight(sheet, headerRow);
        }

        private static int FindFallbackRight(IXLWorksheet sheet, int headerRow)
        {
            var used = sheet.RangeUsed();
            int maxCol = used?.RangeAddress.LastAddress.ColumnNumber ?? sheet.LastColumnUsed().ColumnNumber();
            return maxCol;
        }

        private static int FindSubHeaderRow(IXLWorksheet sheet, int headerRow, int idColLeft)
        {
            int last = sheet.LastRowUsed().RowNumber();
            for (int r = headerRow + 1; r <= Math.Min(headerRow + 6, last); r++)
            {
                bool anyOtm = sheet.Row(r).Cells()
                    .Any(c => c.Address.ColumnNumber != idColLeft &&
                              c.GetString().Trim().StartsWith("Отметка", StringComparison.OrdinalIgnoreCase));
                if (anyOtm) return r;
            }
            return headerRow + 1;
        }

        private static int CountNumericRows(IXLWorksheet sheet, int startRow, int idColLeft, int idColRight)
        {
            int last = sheet.LastRowUsed().RowNumber();
            int count = 0, blanks = 0;

            for (int r = startRow; r <= last; r++)
            {
                var l = sheet.Cell(r, idColLeft).GetString().Trim();
                var rr = idColRight > 0 ? sheet.Cell(r, idColRight).GetString().Trim() : null;

                bool bothEmpty = string.IsNullOrEmpty(l) && string.IsNullOrEmpty(rr);
                if (bothEmpty)
                {
                    blanks++;
                    if (blanks >= 3) break;
                    continue;
                }
                blanks = 0;

                if (int.TryParse(l, out _) && (rr == null || rr == l || int.TryParse(rr, out _)))
                    count++;
                else if (Regex.IsMatch(l, @"^\s*Продолжение", RegexOptions.IgnoreCase))
                    break;
            }
            return count;
        }

        // === Утилита для чтения текста с учётом слитых ячеек ===
        private static string Text(IXLCell cell)
        {
            var s = cell.GetString();
            if (!string.IsNullOrWhiteSpace(s)) return s;
            var mr = cell.MergedRange();
            return mr != null ? mr.FirstCell().GetString() : s;
        }

        // === Поиск стартов циклов (устойчиво к переносу «Цикл № …» на соседнюю строку и merge) ===
        private static List<(int StartColumn, string Label)> FindCycleStartsV2(IXLWorksheet sheet, int headerRow, int subHeaderRow, int idColLeft)
        {
            var starts = new List<(int StartColumn, string Label)>();

            // 1) Ищем подписи «Цикл № …» в окрестности шапки: от headerRow-2 до subHeaderRow+1
            int r1 = Math.Max(1, headerRow - 2);
            int r2 = subHeaderRow + 1;
            for (int r = r1; r <= r2; r++)
            {
                foreach (var c in sheet.Row(r).CellsUsed())
                {
                    var s = Text(c);
                    if (Regex.IsMatch(s ?? string.Empty, @"^\s*Цикл\b", RegexOptions.IgnoreCase))
                    {
                        int candidateCol = FindNearestOtmColumnOnOrRight(sheet, subHeaderRow, c.Address.ColumnNumber, idColLeft);
                        if (candidateCol > 0)
                            starts.Add((candidateCol, s.Trim()));
                    }
                }
            }

            // 2) Добираем все колонки «Отметка …» на строке подзаголовков
            var otmCols = sheet.Row(subHeaderRow).Cells()
                .Where(c => c.Address.ColumnNumber != idColLeft &&
                            c.GetString().Trim().StartsWith("Отметка", StringComparison.OrdinalIgnoreCase))
                .Select(c => c.Address.ColumnNumber)
                .ToList();

            foreach (var col in otmCols)
            {
                if (!starts.Any(s => s.StartColumn == col))
                {
                    string label = FindCycleLabelNear(sheet, headerRow, col);
                    starts.Add((col, label));
                }
            }

            return starts
                .GroupBy(s => s.StartColumn)
                .Select(g => g.First())
                .OrderBy(s => s.StartColumn) // ASC — внутренняя нумерация
                .ToList();
        }

        private static int FindNearestOtmColumnOnOrRight(IXLWorksheet sheet, int subHeaderRow, int fromColumn, int idColLeft)
        {
            int lastCol = sheet.LastColumnUsed().ColumnNumber();
            for (int c = fromColumn; c <= lastCol; c++)
            {
                if (c == idColLeft) continue;
                var s = sheet.Cell(subHeaderRow, c).GetString().Trim();
                if (s.StartsWith("Отметка", StringComparison.OrdinalIgnoreCase))
                    return c;
            }
            for (int c = fromColumn - 1; c >= 1 && c >= fromColumn - 3; c--)
            {
                if (c == idColLeft) continue;
                var s = sheet.Cell(subHeaderRow, c).GetString().Trim();
                if (s.StartsWith("Отметка", StringComparison.OrdinalIgnoreCase))
                    return c;
            }
            return 0;
        }

        private static string FindCycleLabelNear(IXLWorksheet sheet, int headerRow, int aroundColumn)
        {
            // Сканируем несколько строк вокруг headerRow и несколько колонок вокруг искомой
            for (int r = Math.Max(1, headerRow - 2); r <= headerRow + 2; r++)
            {
                for (int dc = -3; dc <= 3; dc++)
                {
                    int c = aroundColumn + dc;
                    if (c <= 0) continue;
                    var s = Text(sheet.Cell(r, c));
                    if (Regex.IsMatch(s ?? string.Empty, @"^\s*Цикл\b", RegexOptions.IgnoreCase))
                        return s.Trim();
                }
            }
            return string.Empty;
        }
    }

    public class WorksheetItem
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
        public IXLWorksheet Sheet { get; set; } = null!;
        public string Display => $"{Index}. {Name}";
    }

    public class ObjectItem
    {
        public int Index { get; set; }
        public int HeaderRow { get; set; }
        public int IdLeftColumn { get; set; }
        public int IdRightColumn { get; set; }
        public int RowsCountHint { get; set; }
        public string Display => RowsCountHint > 0 ? $"{Index}" : $"{Index}";
    }

    public class CycleItem : INotifyPropertyChanged
    {
        public int Index { get; set; }          // внутренняя нумерация цикла (слева-направо)
        public int StartColumn { get; set; }    // Excel-колонка начала тройки
        public string Label { get; set; } = string.Empty;
        public string Display => string.IsNullOrWhiteSpace(Label) ? $"{Index}" : Label;

        public event PropertyChangedEventHandler? PropertyChanged;
    }
}
