using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using System;
using System.Collections.Generic;
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

        private void OkClick(object sender, RoutedEventArgs e)
        {
            var vm = (ImportSelectionVM)DataContext;
            if (!vm.IsValid)
            {
                MessageBox.Show("Выберите лист, объект и цикл.", "Выбор", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            DialogResult = true;
        }

        public IXLWorksheet SelectedWorksheet => ((ImportSelectionVM)DataContext).SelectedWorksheet?.Sheet!;
        public int SelectedObjectIndex => ((ImportSelectionVM)DataContext).SelectedObject?.Index ?? 0;
        public int SelectedCycleIndex => ((ImportSelectionVM)DataContext).SelectedCycle?.Index ?? 0;

        public List<(int Row, IXLCell Cell)> ObjectHeaders => ((ImportSelectionVM)DataContext).CurrentObjectHeaders;
        public List<int> CycleStarts => ((ImportSelectionVM)DataContext).CurrentCycleStarts;
    }

    public partial class ImportSelectionVM : ObservableObject
    {
        private readonly XLWorkbook _wb;
        public ImportSelectionVM(XLWorkbook wb)
        {
            _wb = wb;
            Worksheets = _wb.Worksheets.Select((w, i) => new WorksheetItem { Index = i + 1, Name = w.Name, Sheet = w }).ToList();
            SelectedWorksheet = Worksheets.FirstOrDefault();
        }

        public List<WorksheetItem> Worksheets { get; }
        [ObservableProperty] private WorksheetItem? selectedWorksheet;
        partial void OnSelectedWorksheetChanged(WorksheetItem? value) => LoadObjects();

        public List<ObjectItem> Objects { get; private set; } = new();
        [ObservableProperty] private ObjectItem? selectedObject;
        partial void OnSelectedObjectChanged(ObjectItem? value) => LoadCycles();

        public List<CycleItem> Cycles { get; private set; } = new();
        [ObservableProperty] private CycleItem? selectedCycle;

        public bool IsValid => SelectedWorksheet != null && SelectedObject != null && SelectedCycle != null;

        public List<(int Row, IXLCell Cell)> CurrentObjectHeaders { get; private set; } = new();
        public List<int> CurrentCycleStarts { get; private set; } = new();

        private void LoadObjects()
        {
            Objects.Clear();
            Cycles.Clear();
            SelectedObject = null;
            SelectedCycle = null;
            CurrentObjectHeaders = new();
            CurrentCycleStarts = new();

            if (SelectedWorksheet == null) return;

            var ws = SelectedWorksheet.Sheet;

            CurrentObjectHeaders = FindObjectHeaders(ws);
            Objects = CurrentObjectHeaders
                .Select((h, i) => new ObjectItem
                {
                    Index = i + 1,
                    HeaderRow = h.Row,
                    IdColumn = h.Cell.Address.ColumnNumber
                })
                .ToList();

            OnPropertyChanged(nameof(Objects));
            SelectedObject = Objects.FirstOrDefault();
        }

        private void LoadCycles()
        {
            Cycles.Clear();
            SelectedCycle = null;
            CurrentCycleStarts = new();
            if (SelectedWorksheet == null || SelectedObject == null) return;

            var ws = SelectedWorksheet.Sheet;
            int subHeaderRow = SelectedObject.HeaderRow + 1;
            CurrentCycleStarts = FindCycleStarts(ws, subHeaderRow, SelectedObject.IdColumn);

            var cyclesAsc = CurrentCycleStarts
                               .OrderBy(c => c) 
                               .Select((c, i) => new CycleItem { Index = i + 1, StartColumn = c })
                               .ToList();


            Cycles = cyclesAsc
                        .OrderByDescending(ci => ci.StartColumn)
                        .ToList();
            OnPropertyChanged(nameof(Cycles));
            SelectedCycle = Cycles.FirstOrDefault();
        }

        private static List<(int Row, IXLCell Cell)> FindObjectHeaders(IXLWorksheet sheet) =>
            sheet.RangeUsed()
                 .Rows()
                 .Select(r => (Row: r.RowNumber(),
                               Cell: r.Cells().FirstOrDefault(c =>
                                      Regex.IsMatch(c.GetString(),
                                                    @"^\s*№\s*мар", RegexOptions.IgnoreCase))))
                 .Where(t => t.Cell != null)
                 .ToList();

        private static List<int> FindCycleStarts(IXLWorksheet sheet, int subHdrRow, int idColumn) =>
            sheet.Row(subHdrRow)
                 .CellsUsed()
                 .Where(c =>
                 {
                     string t = c.GetString().Trim();
                     bool isX = Regex.IsMatch(t, @"^\s*[XxХх]\s*$");
                     return isX && c.Address.ColumnNumber != idColumn;
                 })
                 .Select(c => c.Address.ColumnNumber)
                 .ToList();

    }

    public class WorksheetItem
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
        public IXLWorksheet Sheet { get; set; } = default!;
        public string Display => $"{Index}. {Name}";
    }

    public class ObjectItem
    {
        public int Index { get; set; }
        public int HeaderRow { get; set; }
        public int IdColumn { get; set; }
        public string Display => $"{Index}";
    }

    public class CycleItem
    {
        public int Index { get; set; }
        public int StartColumn { get; set; }
        public string Display => $"{Index}";
    }
}
