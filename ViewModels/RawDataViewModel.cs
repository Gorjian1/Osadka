using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using CommunityToolkit.Mvvm.Messaging;
using Microsoft.VisualBasic;
using Osadka.Messages;
using Osadka.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
using System.IO;
namespace Osadka.ViewModels;

public partial class RawDataViewModel : ObservableObject
{
        public ObservableCollection<int> CycleNumbers { get; } = new();
    [ObservableProperty]
    private CycleHeader header = new();
    private readonly Dictionary<int, string> _cycleHeaders = new();
    public IReadOnlyDictionary<int, string> CycleHeaders => _cycleHeaders;

    [ObservableProperty] private string _selectedCycleHeader = string.Empty;

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
        PasteCommand = new RelayCommand(OnPaste);
        LoadFromWorkbookCommand = new RelayCommand(OnLoadWorkbook);
        ClearCommand = new RelayCommand(OnClear);

        Header.PropertyChanged += Header_PropertyChanged;
        WeakReferenceMessenger.Default
            .Register<CoordinatesMessage>(
                this,
                (r, msg) =>
                {
                    CoordRows.Clear();
                    foreach (var pt in msg.Points)
                    {
                        CoordRows.Add(new CoordRow { X = pt.X, Y = pt.Y });
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

                CoordRows.Add(new CoordRow { X = x, Y = y });
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

    public void LoadWorkbookFromFile(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            return;

        try
        {
            using var wb = new XLWorkbook(filePath);

            var ws = SelectWorksheet(wb);
            if (ws == null) return;

            var objHeaders = FindObjectHeaders(ws);
            if (objHeaders.Count == 0)
            {
                MessageBox.Show("На листе не найдены объекты (№ марки).",
                                "Импорт", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int objIdx = (objHeaders.Count == 1)
                ? 1
                : (AskInt($"Объектов: {objHeaders.Count}\nВведите № объекта:",
                          1, objHeaders.Count, out int v) ? v : 0);
            if (objIdx == 0) return;

            var (headerRow, idHeaderCell) = objHeaders[objIdx - 1];
            int idCol = idHeaderCell.Address.ColumnNumber;
            int subHeaderRow = headerRow + 1;

            var cycleStarts = FindCycleStarts(ws, subHeaderRow, idCol);
            if (cycleStarts.Count == 0)
            {
                MessageBox.Show("Не найдено столбцов «Отметка, м».",
                                "Импорт", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int cycleIdx = (cycleStarts.Count == 1)
                ? 1
                : (AskInt($"Циклов: {cycleStarts.Count}\nВведите № цикла:",
                          1, cycleStarts.Count, out int c) ? c : 0);
            if (cycleIdx == 0) return;
            ReadAllObjects(ws, objHeaders, cycleStarts);

            ObjectNumbers.Clear();
            foreach (var k in _objects.Keys.OrderBy(k => k)) ObjectNumbers.Add(k);

            CycleNumbers.Clear();
            foreach (var k in _objects[objIdx].Keys.OrderBy(k => k)) CycleNumbers.Add(k);

            Header.ObjectNumber = objIdx;
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

        void ReadAllObjects(IXLWorksheet sheet,
                            List<(int Row, IXLCell Cell)> headers,
                            List<int> cycleCols)
        {
            _objects.Clear();

            foreach (var (hdr, objNumber) in headers.Select((h, i) => (h, i + 1)))
            {
                int idColLocal = hdr.Cell.Address.ColumnNumber;
                int subHdrRowLocal = hdr.Row + 1;
                int dataRowFirst = subHdrRowLocal + 1;
                int dataRowLast = (objNumber == headers.Count
                                        ? sheet.LastRowUsed().RowNumber()
                                        : headers[objNumber].Row - 1);

                var cyclesDict = new Dictionary<int, List<MeasurementRow>>();

                foreach (var (cycIdx, startCol) in cycleCols.Select((c, i) => (i + 1, c)))
                {
                    var rows = new List<MeasurementRow>();
                    int headerRowLocal = subHdrRowLocal - 1;
                    string cycLabel = sheet.Cell(headerRowLocal, startCol).GetString().Trim();
                    if (!string.IsNullOrWhiteSpace(cycLabel))
                        _cycleHeaders[cycIdx] = cycLabel;

                    for (int r = dataRowFirst; r <= dataRowLast; r++)
                    {
                        var idText = sheet.Cell(r, idColLocal).GetString().Trim();
                        if (string.IsNullOrEmpty(idText)) break;

                        var (mark, markRaw) = ParseCell(sheet.Cell(r, startCol));
                        var (settl, settlRaw) = ParseCell(sheet.Cell(r, startCol + 1));
                        var (total, totalRaw) = ParseCell(sheet.Cell(r, startCol + 2));

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

    private static (int? num, DateTime? date) ParseCycleHeader(string txt)
    {
        int? n = null; DateTime? d = null;

        var m = Regex.Match(txt, @"№\s*(\d+)");
        if (m.Success) n = int.Parse(m.Groups[1].Value);

        m = Regex.Match(txt, @"(\d{2}\.\d{2}\.\d{4})");
        if (m.Success) d = DateTime.ParseExact(
            m.Groups[1].Value, "dd.MM.yyyy", CultureInfo.InvariantCulture);

        return (n, d);
    }


    private static double ReadNumber(IXLCell cell)
    {
        var txt = cell.GetString().Trim();
        if (Regex.IsMatch(txt, @"\bнов", RegexOptions.IgnoreCase))
            return 0;
        if (cell.DataType == XLDataType.Number)
            return cell.GetDouble();
        if (double.TryParse(txt.Replace(',', '.'),
                            NumberStyles.Any,
                            CultureInfo.InvariantCulture,
                            out double val))
            return val;
        return 0;
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
