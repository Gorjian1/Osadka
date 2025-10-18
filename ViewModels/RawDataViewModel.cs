using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using CommunityToolkit.Mvvm.Messaging;
using Microsoft.VisualBasic;
using Microsoft.Win32;
using Osadka.Messages;
using Osadka.Models;
using Osadka.Services;
using Osadka.Core.Units; // Unit, UnitConverter
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

namespace Osadka.ViewModels
{
    public partial class RawDataViewModel : ObservableObject
    {
        private bool _suspendRefresh;
        public void SuspendRefresh(bool on) => _suspendRefresh = on;

        // === Пользовательский шаблон ===
        [ObservableProperty] private string? templatePath;
        public bool HasCustomTemplate => !string.IsNullOrWhiteSpace(TemplatePath) && File.Exists(TemplatePath!);
        partial void OnTemplatePathChanged(string? value)
        {
            UserSettings.Data.TemplatePath = value;
            UserSettings.Save();
            OnPropertyChanged(nameof(HasCustomTemplate));
        }

        [ObservableProperty] private string? drawingPath;

        public class CycleDisplayItem
        {
            public int Number { get; set; }
            public string Label { get; set; } = string.Empty;
        }

        public ObservableCollection<CycleDisplayItem> CycleItems { get; } = new();
        public ObservableCollection<int> CycleNumbers { get; } = new();
        public ObservableCollection<int> ObjectNumbers { get; } = new();
        public ObservableCollection<MeasurementRow> DataRows { get; } = new();
        public ObservableCollection<CoordRow> CoordRows { get; } = new();

        [ObservableProperty] private CycleHeader header = new();
        private readonly Dictionary<int, string> _cycleHeaders = new();
        public IReadOnlyDictionary<int, string> CycleHeaders => _cycleHeaders;

        [ObservableProperty] private string _selectedCycleHeader = string.Empty;

        // === ЕДИНИЦЫ: инвариант → храним ВСЕГДА в мм ===
        // SourceUnit описывает, в чём приходят входные данные (из окна чертежа, буфера и т.д.).
        [ObservableProperty] private Unit sourceUnit = Unit.Millimeter;

        // Для совместимости с существующей разметкой, если она привязана к CoordUnit
        public enum CoordUnits { Millimeters, Centimeters, Decimeters, Meters }
        public IReadOnlyList<CoordUnits> CoordUnitValues { get; } =
            Enum.GetValues(typeof(CoordUnits)).Cast<CoordUnits>().ToList();

        [ObservableProperty]
        private CoordUnits coordUnit = CoordUnits.Millimeters;

        partial void OnCoordUnitChanged(CoordUnits oldVal, CoordUnits newVal)
        {
            // При переключении отображения не трогаем уже сохранённые в мм данные.
            // Держим SourceUnit согласованным с CoordUnit для последующего ввода.
            SourceUnit = Map(newVal);
        }

        // Удобные аксессоры
        public double CoordScale => UnitConverter.ToMm(1.0, Map(coordUnit)); // 1 <ед.> → мм
        private static Unit Map(CoordUnits u) => u switch
        {
            CoordUnits.Millimeters => Unit.Millimeter,
            CoordUnits.Centimeters => Unit.Centimeter,
            CoordUnits.Decimeters => Unit.Decimeter,
            _ => Unit.Meter
        };

        // === Кеш по объектам/циклам ===
        private readonly Dictionary<int, Dictionary<int, List<MeasurementRow>>> _objects = new();
        private readonly Dictionary<int, List<MeasurementRow>> _cycles = new();

        // === Команды ===
        public IRelayCommand OpenTemplate { get; }
        public IRelayCommand ChooseOrOpenTemplateCommand { get; }
        public IRelayCommand ClearTemplateCommand { get; }
        public IRelayCommand PasteCommand { get; }
        public IRelayCommand LoadFromWorkbookCommand { get; }
        public IRelayCommand ClearCommand { get; }

        public RawDataViewModel()
        {
            OpenTemplate = new RelayCommand(OpenTemplatePicker);
            ChooseOrOpenTemplateCommand = new RelayCommand(ChooseOrOpenTemplate);
            ClearTemplateCommand = new RelayCommand(ClearTemplate, () => HasCustomTemplate);

            PasteCommand = new RelayCommand(OnPaste);
            LoadFromWorkbookCommand = new RelayCommand(OnLoadWorkbook);
            ClearCommand = new RelayCommand(OnClear);

            UserSettings.Load();
            TemplatePath = UserSettings.Data.TemplatePath;

            Header.PropertyChanged += Header_PropertyChanged;
            PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(HasCustomTemplate))
                    ((RelayCommand)ClearTemplateCommand).NotifyCanExecuteChanged();
            };

            // ЕДИНСТВЕННОЕ место пересчёта: сразу приводим вход к мм по SourceUnit/CoordUnit
            WeakReferenceMessenger.Default.Register<CoordinatesMessage>(
                this,
                (r, msg) =>
                {
                    CoordRows.Clear();
                    foreach (var pt in msg.Points)
                    {
                        CoordRows.Add(new CoordRow
                        {
                            X = UnitConverter.ToMm(pt.X, Map(coordUnit)),
                            Y = UnitConverter.ToMm(pt.Y, Map(coordUnit))
                        });
                    }
                    OnPropertyChanged(nameof(ShowPlaceholder));
                });
        }

        public bool ShowPlaceholder => DataRows.Count == 0 && CoordRows.Count == 0;

        private void Header_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (_suspendRefresh) return;
            if (e.PropertyName == nameof(Header.CycleNumber) ||
                e.PropertyName == nameof(Header.ObjectNumber))
            {
                RefreshData();
            }
        }


        private void RefreshData()
        {
            CycleNumbers.Clear();
            if (_objects.TryGetValue(Header.ObjectNumber, out var cycles))
            {
                foreach (var k in cycles.Keys.OrderBy(k => k))
                    CycleNumbers.Add(k);

                if (cycles.TryGetValue(Header.CycleNumber, out var rows))
                {
                    DataRows.Clear();
                    foreach (var r in rows) DataRows.Add(r);
                }
            }

            SelectedCycleHeader =
                _cycleHeaders.TryGetValue(Header.CycleNumber, out var cycleHdr)
                    ? cycleHdr
                    : string.Empty;

            // UI список циклов справа
            CycleItems.Clear();
            if (_objects.TryGetValue(Header.ObjectNumber, out var cyclesDict))
            {
                foreach (var k in cyclesDict.Keys.OrderByDescending(k => k))
                {
                    var label = _cycleHeaders.TryGetValue(k, out var h) && !string.IsNullOrWhiteSpace(h)
                                ? h
                                : $"Цикл {k}";
                    CycleItems.Add(new CycleDisplayItem { Number = k, Label = label });
                }
            }

            OnPropertyChanged(nameof(ShowPlaceholder));
        }

        private void OnClear()
        {
            DataRows.Clear();
            CoordRows.Clear();
            _cycles.Clear();
            OnPropertyChanged(nameof(ShowPlaceholder));
        }

        private void ChooseOrOpenTemplate()
        {
            if (HasCustomTemplate)
            {
                try
                {
                    Process.Start(new ProcessStartInfo(TemplatePath!) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Не удалось открыть файл шаблона:\n{ex.Message}", "Шаблон", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            else
            {
                OpenTemplatePicker();
            }
        }

        private void ClearTemplate()
        {
            if (!HasCustomTemplate) return;
            TemplatePath = null;
            MessageBox.Show("Путь к пользовательскому шаблону очищен. Будет использован встроенный template.xlsx.",
                "Шаблон", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // Диалог выбора шаблона
        private void OpenTemplatePicker()
        {
            var dlg = new OpenFileDialog
            {
                Title = "Выберите файл шаблона Excel",
                Filter = "Excel шаблоны (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Все файлы|*.*"
            };
            if (dlg.ShowDialog() == true)
            {
                TemplatePath = dlg.FileName;
                MessageBox.Show("Шаблон успешно выбран:\n" + TemplatePath,
                    "Шаблон", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // === Вставка из буфера ===
        private void OnPaste()
        {
            if (!Clipboard.ContainsText()) return;

            var lines = Clipboard.GetText().Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0) return;
            var first = lines[0].Split('\t');
            int cols = first.Length;

            if (cols == 1)
            {
                int i = 0;
                foreach (var ln in lines)
                {
                    if (i >= DataRows.Count) break;
                    DataRows[i++].Id = ln.Trim();
                }
                return;
            }

            if (cols == 2)
            {
                CoordRows.Clear();
                foreach (var ln in lines)
                {
                    var arr = ln.Split('\t');
                    if (arr.Length < 2) continue;

                    if (!double.TryParse(arr[0].Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double x) ||
                        !double.TryParse(arr[1].Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double y))
                        continue;

                    CoordRows.Add(new CoordRow
                    {
                        X = UnitConverter.ToMm(x, Map(coordUnit)),
                        Y = UnitConverter.ToMm(y, Map(coordUnit))
                    });
                }
                OnPropertyChanged(nameof(ShowPlaceholder));
                return;
            }

            if (cols == 3)
            {
                DataRows.Clear();
                int row = 0;
                foreach (var ln in lines)
                {
                    var arr = ln.Split('\t');
                    if (arr.Length < 3) continue;
                    if (LooksLikeHeader(arr)) continue;

                    var (markVal, markRaw) = TryParse(arr[0]);
                    var (settlVal, settlRaw) = TryParse(arr[1]);
                    var (totalVal, totalRaw) = TryParse(arr[2]);

                    string id = (row < DataRows.Count) ? DataRows[row].Id : (row + 1).ToString();

                    DataRows.Add(new MeasurementRow
                    {
                        Id = id,
                        Mark = markVal,
                        Settl = settlVal,
                        Total = totalVal,
                        MarkRaw = markRaw,
                        SettlRaw = settlRaw,
                        TotalRaw = totalRaw,
                        Cycle = Header.CycleNumber
                    });
                    row++;
                }
                UpdateCache();
                return;
            }

            if (cols == 4)
            {
                DataRows.Clear();
                foreach (var ln in lines)
                {
                    var arr = ln.Split('\t');
                    if (arr.Length < 4) continue;
                    if (LooksLikeHeader(arr)) continue;

                    string markRaw = arr[0];
                    string settlRaw = arr[1];
                    string totalRaw = arr[2];
                    string id = arr[3].Trim();
                    if (string.IsNullOrEmpty(id)) continue;

                    var (markVal, _) = TryParse(markRaw);
                    var (settlVal, _) = TryParse(settlRaw);
                    var (totalVal, _) = TryParse(totalRaw);

                    DataRows.Add(new MeasurementRow
                    {
                        Id = id,
                        Mark = markVal,
                        Settl = settlVal,
                        Total = totalVal,
                        MarkRaw = markRaw,
                        SettlRaw = settlRaw,
                        TotalRaw = totalRaw,
                        Cycle = Header.CycleNumber
                    });
                }
                UpdateCache();
                return;
            }

            OnPropertyChanged(nameof(DataRows));
            MessageBox.Show("Формат буфера не поддерживается (должно быть 1-4 колонок).");
        }

        private static (double? val, string raw) TryParse(string txt)
        {
            txt = txt.Trim();
            if (Regex.IsMatch(txt, @"\bнов", RegexOptions.IgnoreCase))
                return (0, txt);

            if (double.TryParse(txt.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double v))
                return (v, txt);

            return (null, txt);
        }

        private static bool LooksLikeHeader(string[] cells)
        {
            var joined = string.Join(" ", cells).ToLowerInvariant();
            if (joined.Contains("отмет") || joined.Contains("осад") || joined.Contains("суммар") ||
                joined.Contains("№") || joined.Contains("марка") || joined.Contains("cycle") || joined.Contains("id"))
                return true;

            int nonNumeric = 0;
            for (int i = 0; i < cells.Length; i++)
            {
                var t = (cells[i] ?? string.Empty).Trim();
                if (Regex.IsMatch(t, @"\bнов", RegexOptions.IgnoreCase))
                    continue;

                if (!double.TryParse(t.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                    nonNumeric++;
            }
            return nonNumeric >= Math.Max(2, cells.Length - 1);
        }

        // === Импорт из Excel ===
        private void OnLoadWorkbook()
        {
            var dlg = new OpenFileDialog { Filter = "Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm" };
            if (dlg.ShowDialog() == true)
                LoadWorkbookFromFile(dlg.FileName);
        }

        public void LoadWorkbookFromFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath)) return;

            try
            {
                using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                using var wb = new XLWorkbook(stream);

                var dlg = new Osadka.Views.ImportSelectionWindow(wb)
                {
                    Owner = Application.Current?.MainWindow
                };
                if (dlg.ShowDialog() != true) return;

                IXLWorksheet ws = dlg.SelectedWorksheet?.Sheet ?? throw new InvalidOperationException("Не выбран лист Excel.");

                var objHeaders = dlg.ObjectHeaders;
                var cycleStarts = dlg.CycleStarts;
                int objIdx = dlg.SelectedObjectIndex;   // 1-based
                int cycleIdx = dlg.SelectedCycleIndex;  // 1-based

                if (objHeaders == null || objHeaders.Count == 0) objHeaders = FindObjectHeaders(ws);
                if (objHeaders == null || objHeaders.Count == 0)
                    throw new InvalidOperationException("Не удалось найти заголовок с «№ точки» на листе.");

                var hdrTuple = objIdx >= 1 && objIdx <= objHeaders.Count ? objHeaders[objIdx - 1] : objHeaders.First();
                int idCol = hdrTuple.Cell.Address.ColumnNumber;
                int subHdrRow = FindSubHeaderRow(ws, hdrTuple.Row, idCol);

                if (cycleStarts == null || cycleStarts.Count == 0)
                {
                    var computed = FindCycleStarts(ws, subHdrRow, idCol);
                    if (computed.Count == 0)
                    {
                        int lastRow = ws.LastRowUsed().RowNumber();
                        for (int r = hdrTuple.Row; r <= Math.Min(hdrTuple.Row + 10, lastRow); r++)
                        {
                            bool anyOtm = ws.Row(r).Cells().Any(c => Regex.IsMatch(c.GetString(), @"^\s*Отметка", RegexOptions.IgnoreCase));
                            if (anyOtm)
                            {
                                subHdrRow = r;
                                computed = FindCycleStarts(ws, subHdrRow, idCol);
                                if (computed.Count > 0) break;
                            }
                        }
                    }
                    cycleStarts = computed;
                }

                _cycleHeaders.Clear();
                ReadAllObjects(ws, objHeaders, cycleStarts);

                ObjectNumbers.Clear();
                foreach (var k in _objects.Keys.OrderBy(k => k)) ObjectNumbers.Add(k);

                Header.ObjectNumber = (objIdx >= 1 && objIdx <= ObjectNumbers.Count)
                    ? ObjectNumbers[objIdx - 1]
                    : (ObjectNumbers.Count > 0 ? ObjectNumbers[0] : 1);

                CycleNumbers.Clear();
                if (_objects.TryGetValue(Header.ObjectNumber, out var cyclesForObject))
                {
                    foreach (var k in cyclesForObject.Keys.OrderBy(k => k)) CycleNumbers.Add(k);
                }

                if (CycleNumbers.Count > 0)
                {
                    int idx = Math.Clamp(cycleIdx, 1, CycleNumbers.Count);
                    int chosenNumber = CycleNumbers[idx - 1];
                    Header.CycleNumber = chosenNumber;
                }

                RefreshData();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при импорте Excel:\n{ex.Message}", "Импорт", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            // === Локальные функции ===
            List<(int Row, IXLCell Cell)> FindObjectHeaders(IXLWorksheet sheet)
                => sheet.RangeUsed()?
                       .Rows()
                       .Select(r =>
                       {
                           var hits = r.Cells().Where(c => Regex.IsMatch(c.GetString(), @"^\s*№\s*(точки|мар\w*)", RegexOptions.IgnoreCase));
                           if (!hits.Any()) return (Row: 0, Cell: (IXLCell?)null);
                           var leftMost = hits.OrderBy(c => c.Address.ColumnNumber).First();
                           return (Row: r.RowNumber(), Cell: leftMost);
                       })
                       .Where(t => t.Cell != null && t.Row > 0)
                       .ToList()
                   ?? new List<(int Row, IXLCell Cell)>();

            List<int> FindCycleStarts(IXLWorksheet sheet, int subHdrRow, int idColumn)
                => sheet.Row(subHdrRow)
                        .Cells()
                        .Where(c => c.Address.ColumnNumber != idColumn && c.GetString().Trim().StartsWith("Отметка", StringComparison.OrdinalIgnoreCase))
                        .Select(c => c.Address.ColumnNumber)
                        .Distinct()
                        .OrderBy(c => c)
                        .ToList();

            int FindSubHeaderRow(IXLWorksheet s, int headerRow, int idColumn)
            {
                int lastRow = s.LastRowUsed().RowNumber();
                for (int r = headerRow + 1; r <= Math.Min(headerRow + 6, lastRow); r++)
                {
                    bool ok = s.Row(r).Cells().Any(c => c.Address.ColumnNumber != idColumn && c.GetString().Trim().StartsWith("Отметка", StringComparison.OrdinalIgnoreCase));
                    if (ok) return r;
                }
                return headerRow + 1;
            }

            void ReadAllObjects(IXLWorksheet sheet, List<(int Row, IXLCell Cell)> headers, List<int> cycleCols)
            {
                _objects.Clear();
                if (headers == null || headers.Count == 0) return;

                headers = headers.OrderBy(h => h.Row).ToList();

                for (int objNumber = 1; objNumber <= headers.Count; objNumber++)
                {
                    var hdr = headers[objNumber - 1];

                    int idColLocal = hdr.Cell.Address.ColumnNumber;
                    int subHdrRowLocal = FindSubHeaderRow(sheet, hdr.Row, idColLocal);

                    int dataRowFirst = subHdrRowLocal + 1;
                    int dataRowLast = (objNumber == headers.Count ? sheet.LastRowUsed().RowNumber() : headers[objNumber].Row - 1);

                    var localCycCols = (cycleCols != null && cycleCols.Count > 0) ? cycleCols : FindCycleStarts(sheet, subHdrRowLocal, idColLocal);

                    var cyclesDict = new Dictionary<int, List<MeasurementRow>>();

                    foreach (var (cycIdx, startCol) in localCycCols.Select((c, i) => (i + 1, c)))
                    {
                        string cycLabel = BuildCycleHeaderLabel(sheet, startCol, subHdrRowLocal, hdr.Row);
                        if (!string.IsNullOrWhiteSpace(cycLabel)) _cycleHeaders[cycIdx] = cycLabel;

                        var rows = new List<MeasurementRow>();
                        int blanksInARow = 0;

                        for (int r = dataRowFirst; r <= dataRowLast; r++)
                        {
                            string idText = sheet.Cell(r, idColLocal).GetString().Trim();
                            if (Regex.IsMatch(idText, @"^\s*№\s*(точки|мар\w*)", RegexOptions.IgnoreCase)) break;

                            if (string.IsNullOrEmpty(idText))
                            {
                                blanksInARow++;
                                if (blanksInARow >= 3) break;
                                continue;
                            }
                            blanksInARow = 0;

                            var (mark, markRaw) = ParseCell(sheet.Cell(r, startCol));
                            var (settl, settlRaw) = ParseCell(sheet.Cell(r, startCol + 1));
                            var (total, totalRaw) = ParseCell(sheet.Cell(r, startCol + 2));
                            if (mark.HasValue) mark = UnitConverter.ToMm(mark.Value, Map(coordUnit));
                            if (settl.HasValue) settl = UnitConverter.ToMm(settl.Value, Map(coordUnit));
                            if (total.HasValue) total = UnitConverter.ToMm(total.Value, Map(coordUnit));
                            if (mark is null && settl is null && total is null &&
                                string.IsNullOrWhiteSpace(markRaw) && string.IsNullOrWhiteSpace(settlRaw) && string.IsNullOrWhiteSpace(totalRaw))
                            {
                                continue;
                            }
                            if (settl.HasValue) settl = Math.Round(settl.Value, 1);
                            if (total.HasValue) total = Math.Round(total.Value, 1);

                            rows.Add(new MeasurementRow
                            {
                                Id = idText,
                                Mark = mark,
                                Settl = settl,
                                Total = total,
                                MarkRaw = markRaw,
                                SettlRaw = settlRaw,
                                TotalRaw = totalRaw
                            });
                        }

                        cyclesDict[cycIdx] = rows;
                    }

                    _objects[objNumber] = cyclesDict;
                }
            }

            string BuildCycleHeaderLabel(IXLWorksheet sheet, int startCol, int subHdrRow, int headerRow)
            {
                string Read(IXLCell cell)
                {
                    var s = cell.GetString();
                    if (!string.IsNullOrWhiteSpace(s)) return s;
                    var mr = cell.MergedRange();
                    return mr != null ? mr.FirstCell().GetString() : s;
                }

                int r1 = Math.Max(1, headerRow - 2);
                int r2 = subHdrRow + 1;

                // 1) Ищем только внутри текущей тройки (Отметка/Осадка/Общая)
                for (int r = r1; r <= r2; r++)
                {
                    for (int c = startCol; c <= startCol + 2; c++)
                    {
                        var s = Read(sheet.Cell(r, c));
                        if (!string.IsNullOrWhiteSpace(s) && Regex.IsMatch(s, @"^\s*Цикл\b", RegexOptions.IgnoreCase))
                            return s.Trim();
                    }
                }

                // 2) Фолбэк — центр-сначала (0,+1,-1,+2,-2,...)
                int[] offs = new[] { 0, +1, -1, +2, -2, +3, -3 };
                for (int r = r1; r <= r2; r++)
                {
                    foreach (var dc in offs)
                    {
                        int c = startCol + dc;
                        if (c <= 0) continue;
                        var s = Read(sheet.Cell(r, c));
                        if (!string.IsNullOrWhiteSpace(s) && Regex.IsMatch(s, @"^\s*Цикл\b", RegexOptions.IgnoreCase))
                            return s.Trim();
                    }
                }

                return string.Empty;
            }
        }

        private static (double? val, string raw) ParseCell(IXLCell cell)
        {
            string txt = cell.GetString().Trim();
            if (Regex.IsMatch(txt, @"\bнов", RegexOptions.IgnoreCase)) return (0, txt);

            if (cell.DataType == XLDataType.Number) return (cell.GetDouble(), txt);

            if (double.TryParse(txt.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double v))
                return (v, txt);

            return (null, txt);
        }

        private void UpdateCache()
        {
            if (!_objects.TryGetValue(Header.ObjectNumber, out var cycles))
            {
                cycles = new Dictionary<int, List<MeasurementRow>>();
                _objects[Header.ObjectNumber] = cycles;
            }
            cycles[Header.CycleNumber] = DataRows.ToList();
            OnPropertyChanged(nameof(ShowPlaceholder));
        }

        public Dictionary<int, Dictionary<int, List<MeasurementRow>>> Objects => _objects;
        public IReadOnlyDictionary<int, List<MeasurementRow>> CurrentCycles =>
            _objects.TryGetValue(Header.ObjectNumber, out var cycles)
                ? cycles
                : new Dictionary<int, List<MeasurementRow>>();

        // Небольшая утилита-вопрос для некоторых сценариев импорта
        private static bool AskInt(string prompt, int min, int max, out int value)
        {
            value = 0;
            string s = Interaction.InputBox(prompt, "Выбор", min.ToString());
            return int.TryParse(s, out value) && value >= min && value <= max;
        }
    }
}
