using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using CommunityToolkit.Mvvm.Messaging;
using Microsoft.VisualBasic;
using Microsoft.Win32;
using Osadka.Messages;
using Osadka.Models;
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
using Osadka.Services;
using System.Windows; // для MessageBox
namespace Osadka.ViewModels;

public partial class RawDataViewModel : ObservableObject
{
    // === Новый путь к пользовательскому шаблону ===
    [ObservableProperty] private string? templatePath;
    public bool HasCustomTemplate => !string.IsNullOrWhiteSpace(TemplatePath) && File.Exists(TemplatePath!);
    partial void OnTemplatePathChanged(string? value)
    {
        UserSettings.Data.TemplatePath = value;
        UserSettings.Save();

        OnPropertyChanged(nameof(HasCustomTemplate));
    }


    public IRelayCommand OpenTemplate { get; }
    public IRelayCommand ChooseOrOpenTemplateCommand { get; }  // ← НОВАЯ (левая часть сплит-кнопки)
    public IRelayCommand ClearTemplateCommand { get; }
    public ObservableCollection<int> CycleNumbers { get; } = new();
    [ObservableProperty]
    private CycleHeader header = new();
    private readonly Dictionary<int, string> _cycleHeaders = new();
    public IReadOnlyDictionary<int, string> CycleHeaders => _cycleHeaders;

    [ObservableProperty] private string _selectedCycleHeader = string.Empty;

    public enum CoordUnits
    {
        Millimeters,
        Centimeters,
        Decimeters,
        Meters
    }

    public IReadOnlyList<CoordUnits> CoordUnitValues { get; } =
        Enum.GetValues(typeof(CoordUnits)).Cast<CoordUnits>().ToList();

    [ObservableProperty]
    private CoordUnits coordUnit = CoordUnits.Millimeters;

    public double CoordScale => coordUnit switch
    {
        CoordUnits.Millimeters => 1,
        CoordUnits.Centimeters => 10,
        CoordUnits.Decimeters => 100,
        _ => 1000        // Meters
    };

    partial void OnCoordUnitChanged(CoordUnits oldVal, CoordUnits newVal)
    {
        double oldScale = oldVal switch
        {
            CoordUnits.Millimeters => 1,
            CoordUnits.Centimeters => 10,
            CoordUnits.Decimeters => 100,
            _ => 1000        // Meters
        };

        double k = CoordScale / oldScale;

        foreach (var p in CoordRows)
        {
            p.X *= k;
            p.Y *= k;
        }
    }

    private void RefreshData()
    {
        CycleNumbers.Clear();
        if (_objects.TryGetValue(Header.ObjectNumber, out var cycles))
        {
            foreach (var k in cycles.Keys.OrderBy(k => k)) CycleNumbers.Add(k);

            if (cycles.TryGetValue(Header.CycleNumber, out var rows))
            {
                DataRows.Clear();
                foreach (var r in rows) DataRows.Add(r);
            }
        }

        SelectedCycleHeader =
            _cycleHeaders.TryGetValue(Header.CycleNumber, out var h) ? h : string.Empty;

        OnPropertyChanged(nameof(ShowPlaceholder));
    }

    private readonly Dictionary<int, Dictionary<int, List<MeasurementRow>>> _objects = new();
    public ObservableCollection<int> ObjectNumbers { get; } = new();
    public ObservableCollection<MeasurementRow> DataRows { get; } = new();
    public ObservableCollection<CoordRow> CoordRows { get; } = new();

    private readonly Dictionary<int, List<MeasurementRow>> _cycles = new();
    public IRelayCommand PasteCommand { get; }
    public IRelayCommand LoadFromWorkbookCommand { get; }
    public IRelayCommand ClearCommand { get; }

    public RawDataViewModel()
    {
        OpenTemplate = new RelayCommand(Opentemp);
        PasteCommand = new RelayCommand(OnPaste);
        LoadFromWorkbookCommand = new RelayCommand(OnLoadWorkbook);
        ClearCommand = new RelayCommand(OnClear);
        UserSettings.Load();
        TemplatePath = UserSettings.Data.TemplatePath;
        Header.PropertyChanged += Header_PropertyChanged;
        ChooseOrOpenTemplateCommand = new RelayCommand(ChooseOrOpenTemplate);
        ClearTemplateCommand = new RelayCommand(ClearTemplate, () => HasCustomTemplate);
        PropertyChanged += (_, e) =>
                {
                        if (e.PropertyName == nameof(HasCustomTemplate))
                ((RelayCommand)ClearTemplateCommand).NotifyCanExecuteChanged();
                   }
        ;
        WeakReferenceMessenger.Default
            .Register<CoordinatesMessage>(
                this,
                (r, msg) =>
                {
                    CoordRows.Clear();
                    foreach (var pt in msg.Points)
                    {
                        CoordRows.Add(new CoordRow { X = pt.X * CoordScale, Y = pt.Y * CoordScale });
                    }
                    OnPropertyChanged(nameof(ShowPlaceholder));
                });
    }

    public bool ShowPlaceholder => DataRows.Count == 0 && CoordRows.Count == 0;

    private void Header_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(Header.CycleNumber) ||
            e.PropertyName == nameof(Header.ObjectNumber))
        {
            RefreshData();
        }
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
                MessageBox.Show($"Не удалось открыть файл шаблона:\n{ex.Message}",
                                "Шаблон", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        else
        {
            Opentemp(); // просто диалог выбора + уведомление
        }
    }

    // Правая часть сплит-кнопки: очистить путь
    private void ClearTemplate()
    {
        if (!HasCustomTemplate) return;
        TemplatePath = null; // сработает OnTemplatePathChanged → сохраняем настройки
        MessageBox.Show("Путь к пользовательскому шаблону очищен. Будет использован встроенный template.xlsx.",
        "Шаблон", MessageBoxButton.OK, MessageBoxImage.Information);
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
            if (System.Text.RegularExpressions.Regex.IsMatch(t, @"\bнов", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                continue;

            if (!double.TryParse(t.Replace(',', '.'), System.Globalization.NumberStyles.Any,
                                  System.Globalization.CultureInfo.InvariantCulture, out _))
                nonNumeric++;
        }
        return nonNumeric >= Math.Max(2, cells.Length - 1);
    }
    private void OnPaste()
    {
        if (!Clipboard.ContainsText()) return;

        var lines = Clipboard.GetText()
                             .Split(new[] { '\r', '\n' },
                                    StringSplitOptions.RemoveEmptyEntries);
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

                if (!double.TryParse(arr[0].Replace(',', '.'),
                                     NumberStyles.Any, CultureInfo.InvariantCulture, out double x) ||
                    !double.TryParse(arr[1].Replace(',', '.'),
                                     NumberStyles.Any, CultureInfo.InvariantCulture, out double y))
                    continue;

                CoordRows.Add(new CoordRow { X = x * CoordScale, Y = y * CoordScale });
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

    private static (double? val, string raw) TryParse(string txt)
    {
        txt = txt.Trim();

        if (Regex.IsMatch(txt, @"\bнов", RegexOptions.IgnoreCase))
            return (0, txt);

        if (double.TryParse(txt.Replace(',', '.'),
                            NumberStyles.Any,
                            CultureInfo.InvariantCulture,
                            out double v))
            return (v, txt);

        return (null, txt);
    }

    private void OnLoadWorkbook()
    {
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm"
        };
        if (dlg.ShowDialog() == true)
            LoadWorkbookFromFile(dlg.FileName);
    }

    // === ПЕРЕПИСАНО: выбор пользовательского шаблона ===
    private void Opentemp()
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
        // Отмена — ничего не меняем. Если TemplatePath пуст, при экспорте возьмётся встроенный.
    }

    public void LoadWorkbookFromFile(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            return;

        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var wb = new XLWorkbook(stream);

            var dlg = new Osadka.Views.ImportSelectionWindow(wb)
            {
                Owner = Application.Current?.MainWindow
            };

            if (dlg.ShowDialog() != true)
                return;

            var ws = dlg.SelectedWorksheet;
            var objHeaders = dlg.ObjectHeaders;
            var cycleStarts = dlg.CycleStarts;
            int objIdx = dlg.SelectedObjectIndex;
            int cycleIdx = dlg.SelectedCycleIndex;

            if (cycleStarts == null || cycleStarts.Count == 0)
            {
                if (objHeaders == null || objHeaders.Count == 0)
                    throw new InvalidOperationException("Не найдены заголовки объектов на листе.");

                var hdrTuple = objIdx >= 1 && objIdx <= objHeaders.Count
                    ? objHeaders[objIdx - 1]
                    : objHeaders.First();

                int idCol = hdrTuple.Cell.Address.ColumnNumber;
                int subHdrRow = FindSubHeaderRow(ws, hdrTuple.Row, idCol, new List<int>());
                var computed = FindCycleStarts(ws, subHdrRow, idCol);

                if (computed.Count == 0)
                {
                    int lastRow = ws.LastRowUsed().RowNumber();
                    for (int r = hdrTuple.Row; r <= Math.Min(hdrTuple.Row + 10, lastRow); r++)
                    {
                        var anyOtm = ws.Row(r).Cells().Any(c =>
                            Regex.IsMatch(c.GetString(), @"^\s*Отметка", RegexOptions.IgnoreCase));
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

            CycleNumbers.Clear();
            if (_objects.TryGetValue(objIdx, out var dict))
            {
                foreach (var k in dict.Keys.OrderBy(k => k)) CycleNumbers.Add(k);
            }

            Header.ObjectNumber = objIdx;

            if (CycleNumbers.Count > 0 && !CycleNumbers.Contains(cycleIdx))
                Header.CycleNumber = CycleNumbers[0];
            else
                Header.CycleNumber = cycleIdx;

            RefreshData();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка при импорте Excel:\n{ex.Message}",
                            "Импорт", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        IXLWorksheet? SelectWorksheet(XLWorkbook book)
        {
            if (book.Worksheets.Count == 0) return null;
            if (book.Worksheets.Count == 1) return book.Worksheet(1);

            var list = book.Worksheets.Select((w, i) => $"{i + 1}. {w.Name}").ToArray();
            return AskInt("Доступные листы:\n" + string.Join("\n", list) +
                          "\nВведите № листа:", 1, list.Length, out int idx)
                   ? book.Worksheet(idx)
                   : null;
        }

        List<(int Row, IXLCell Cell)> FindObjectHeaders(IXLWorksheet sheet) =>
            sheet.RangeUsed()
                 .Rows()
                 .Select(r => (Row: r.RowNumber(),
                               Cell: r.Cells().FirstOrDefault(c =>
                                      Regex.IsMatch(c.GetString(),
                                                    @"^\s*№\s*мар", RegexOptions.IgnoreCase))))
                 .Where(t => t.Cell != null)
                 .ToList();

        List<int> FindCycleStarts(IXLWorksheet sheet, int subHdrRow, int idColumn) =>
            sheet.Row(subHdrRow)
                 .Cells()
                 .Where(c => c.GetString().Trim()
                              .StartsWith("Отметка", StringComparison.OrdinalIgnoreCase))
                 .Select(c => c.Address.ColumnNumber)
                 .Where(col => col != idColumn)
                 .ToList();

        int FindSubHeaderRow(IXLWorksheet s, int headerRow, int idColumn, List<int> cycCols)
        {
            int lastRow = s.LastRowUsed().RowNumber();
            int row = headerRow + 1;
            int maxLookahead = Math.Min(headerRow + 6, lastRow);

            for (; row <= maxLookahead; row++)
            {
                bool looksLikeSub = cycCols.Any(c =>
                    Regex.IsMatch(s.Cell(row, c).GetString(), @"отметка", RegexOptions.IgnoreCase));

                bool nextHasOtm = row + 1 <= lastRow && cycCols.Any(c =>
                    Regex.IsMatch(s.Cell(row + 1, c).GetString(), @"отметка", RegexOptions.IgnoreCase));
                bool thisHasAny = cycCols.Any(c => !string.IsNullOrWhiteSpace(s.Cell(row, c).GetString()));

                if (looksLikeSub) return row;
                if (thisHasAny && nextHasOtm) return row + 1;
            }

            return headerRow + 1;
        }

        string BuildCycleHeaderLabel(IXLWorksheet s, int startCol, int subRow, int headerRow)
        {
            var parts = new List<string>();
            int r1 = subRow - 1;
            int r2 = subRow - 2;

            if (r2 >= headerRow && !string.IsNullOrWhiteSpace(s.Cell(r2, startCol).GetString()))
                parts.Add(s.Cell(r2, startCol).GetString().Trim());

            if (r1 >= headerRow && !string.IsNullOrWhiteSpace(s.Cell(r1, startCol).GetString()))
                parts.Add(s.Cell(r1, startCol).GetString().Trim());

            return string.Join(" ", parts.Where(p => !string.IsNullOrWhiteSpace(p)));
        }

        void ReadAllObjects(IXLWorksheet sheet,
                            List<(int Row, IXLCell Cell)> headers,
                            List<int> cycleCols)
        {
            _objects.Clear();

            foreach (var (hdr, objNumber) in headers.Select((h, i) => (h, i + 1)))
            {
                int idColLocal = hdr.Cell.Address.ColumnNumber;

                int subHdrRowLocal = FindSubHeaderRow(sheet, hdr.Row, idColLocal, cycleCols);

                int dataRowFirst = subHdrRowLocal + 1;
                int dataRowLast = (objNumber == headers.Count
                                        ? sheet.LastRowUsed().RowNumber()
                                        : headers[objNumber].Row - 1);

                var cyclesDict = new Dictionary<int, List<MeasurementRow>>();

                foreach (var (cycIdx, startCol) in cycleCols.Select((c, i) => (i + 1, c)))
                {
                    string cycLabel = BuildCycleHeaderLabel(sheet, startCol, subHdrRowLocal, hdr.Row);
                    if (!string.IsNullOrWhiteSpace(cycLabel))
                        _cycleHeaders[cycIdx] = cycLabel;

                    var rows = new List<MeasurementRow>();
                    int blanksInARow = 0;

                    for (int r = dataRowFirst; r <= dataRowLast; r++)
                    {
                        string idText = sheet.Cell(r, idColLocal).GetString().Trim();

                        if (Regex.IsMatch(idText, @"^\s*№\s*мар", RegexOptions.IgnoreCase))
                            break;

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

                        if (mark is null && settl is null && total is null &&
                            string.IsNullOrWhiteSpace(markRaw) &&
                            string.IsNullOrWhiteSpace(settlRaw) &&
                            string.IsNullOrWhiteSpace(totalRaw))
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
                            TotalRaw = totalRaw,
                            Cycle = cycIdx
                        });
                    }

                    cyclesDict[cycIdx] = rows;
                }

                _objects[objNumber] = cyclesDict;
            }
        }
    }

    private static (double? val, string raw) ParseCell(IXLCell cell)
    {
        string txt = cell.GetString().Trim();

        if (Regex.IsMatch(txt, @"\bнов", RegexOptions.IgnoreCase))
            return (0, txt);

        if (cell.DataType == XLDataType.Number)
            return (cell.GetDouble(), txt);

        if (double.TryParse(txt.Replace(',', '.'),
                            NumberStyles.Any,
                            CultureInfo.InvariantCulture,
                            out double v))
            return (v, txt);

        return (null, txt);
    }

    public Dictionary<int, Dictionary<int, List<MeasurementRow>>> Objects
       => _objects;

    public IReadOnlyDictionary<int, List<MeasurementRow>> CurrentCycles =>
        _objects.TryGetValue(Header.ObjectNumber, out var cycles)
            ? cycles
            : new Dictionary<int, List<MeasurementRow>>();

    private static bool AskInt(string prompt, int min, int max, out int value)
    {
        value = 0;
        string s = Interaction.InputBox(prompt, "Выбор", min.ToString());
        return int.TryParse(s, out value) && value >= min && value <= max;
    }
}
