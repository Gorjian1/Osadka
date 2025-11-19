// File: ViewModels/MainViewModel.cs
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Windows;
using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using Osadka.Models;
using Osadka.Services;
using Osadka.Services.Abstractions;
using Osadka.Views;

namespace Osadka.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private string? _currentPath;

        // Services (injected via DI)
        private readonly IMessageBoxService _messageBox;
        private readonly IFileDialogService _fileDialog;
        private readonly IFileService _fileService;
        private readonly IProjectService _projectService;
        private readonly IExportService _exportService;
        private readonly INavigationService _navigationService;

        public RawDataViewModel RawVM { get; }
        public GeneralReportViewModel GenVM { get; }
        public RelativeSettlementsViewModel RelVM { get; }
        public DynamicsGrafficViewModel DynVM { get; }
        private readonly DynamicsReportService _dynSvc;
        public IRelayCommand HelpCommand { get; }
        public IRelayCommand<string> NavigateCommand { get; }
        public IRelayCommand NewProjectCommand { get; }
        public IRelayCommand OpenProjectCommand { get; }
        public IRelayCommand SaveProjectCommand { get; }
        public IRelayCommand SaveAsProjectCommand { get; }
        public IRelayCommand QuickReportCommand { get; }

        private readonly HashSet<MeasurementRow> _trackedMeasurementRows = new();
        private readonly HashSet<CoordRow> _trackedCoordRows = new();
        private bool _isDirty;
        private bool _includeGeneral = true;
        public bool IncludeGeneral
        {
            get => _includeGeneral;
            set => SetProperty(ref _includeGeneral, value);
        }

        private bool _includeRelative = true;
        public bool IncludeRelative
        {
            get => _includeRelative;
            set => SetProperty(ref _includeRelative, value);
        }

        private bool _includeGraphs = true;
        public bool IncludeGraphs
        {
            get => _includeGraphs;
            set => SetProperty(ref _includeGraphs, value);
        }

        public object? CurrentPage => _navigationService.CurrentPage;

        private static class PageKeys
        {
            public const string Raw = "RawData";
            public const string Diff = "SettlementDiff";
            public const string Sum = "Summary";
            public const string Graf = "Graffics";
            public const string Coord = "Coordinates";
        }

        private void OpenHelp()
        {
            string exeDir = AppContext.BaseDirectory;
            string docx = Path.Combine(exeDir, "help.docx");

            if (_fileService.FileExists(docx))
                _fileService.OpenInDefaultApp(docx);
            else
                _messageBox.Show("Файл справки не найден.", "Справка");
        }

        public MainViewModel(
            RawDataViewModel rawDataViewModel,
            IMessageBoxService messageBox,
            IFileDialogService fileDialog,
            IFileService fileService,
            IProjectService projectService,
            IExportService exportService,
            INavigationService navigationService,
            GeneralReportService generalReportService,
            RelativeReportService relativeReportService,
            DynamicsReportService dynamicsReportService)
        {
            // Inject services
            _messageBox = messageBox;
            _fileDialog = fileDialog;
            _fileService = fileService;
            _projectService = projectService;
            _exportService = exportService;
            _navigationService = navigationService;

            // Inject ViewModels and Services
            RawVM = rawDataViewModel;
            _dynSvc = dynamicsReportService;

            // Create other ViewModels
            GenVM = new GeneralReportViewModel(RawVM, generalReportService, relativeReportService);
            RelVM = new RelativeSettlementsViewModel(RawVM, relativeReportService);

            // Register navigation pages
            RegisterPages();

            // Setup event subscriptions
            RawVM.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(RawDataViewModel.TemplatePath) ||
                    e.PropertyName == nameof(RawDataViewModel.CoordUnit) ||
                    e.PropertyName == nameof(RawDataViewModel.DrawingPath))
                {
                    MarkDirty();
                }
            };

            RawVM.Header.PropertyChanged += Header_PropertyChanged;
            RawVM.DataRows.CollectionChanged += DataRows_CollectionChanged;
            RawVM.CoordRows.CollectionChanged += CoordRows_CollectionChanged;

            foreach (var row in RawVM.DataRows)
                SubscribeMeasurementRow(row);
            foreach (var row in RawVM.CoordRows)
                SubscribeCoordRow(row);

            // Setup commands
            HelpCommand = new RelayCommand(OpenHelp);
            NavigateCommand = new RelayCommand<string>(Navigate);
            NewProjectCommand = new RelayCommand(NewProject);
            OpenProjectCommand = new RelayCommand(OpenProject);
            SaveProjectCommand = new RelayCommand(SaveProject, () => _isDirty);
            SaveAsProjectCommand = new RelayCommand(SaveAsProject);
            QuickReportCommand = new RelayCommand(DoQuickExport, () => GenVM.Report != null);

            GenVM.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(GeneralReportViewModel.Report) &&
                    QuickReportCommand is RelayCommand quickRelay)
                {
                    quickRelay.NotifyCanExecuteChanged();
                }
            };

            // Subscribe to navigation changes to notify UI
            _navigationService.CurrentPageChanged += (_, _) => OnPropertyChanged(nameof(CurrentPage));

            // Navigate to initial page
            Navigate(PageKeys.Raw);
        }

        #region Navigation

        private void RegisterPages()
        {
            // Register all navigation pages with their factories
            _navigationService.RegisterPage(PageKeys.Raw, () => new RawDataPage(RawVM));
            _navigationService.RegisterPage(PageKeys.Diff, () => new GeneralReportPage(GenVM));
            _navigationService.RegisterPage(PageKeys.Sum, () => new RelativeSettlementsPage(RelVM));
            _navigationService.RegisterPage(PageKeys.Coord, () => new CoordinateExporting(RawVM));
            _navigationService.RegisterPage(PageKeys.Graf, () =>
                new DynamicsGrafficPage(new DynamicsGrafficViewModel(RawVM, _dynSvc)));
        }

        private void Navigate(string? key)
        {
            if (string.IsNullOrWhiteSpace(key)) return;
            _navigationService.NavigateTo(key);
        }

        #endregion

        private void NewProject()
        {
            RawVM.ClearCommand.Execute(null);
            _currentPath = null;
            ResetDirty();
        }

        public void LoadProject(string path)
        {
            if (RawVM is not { } vm) return;

            try
            {
                var (data, dwgPath) = _projectService.Load(path);

                vm.Header.CycleNumber = data.Cycle;
                vm.Header.MaxNomen = data.MaxNomen;
                vm.Header.MaxCalculated = data.MaxCalculated;
                vm.Header.RelNomen = data.RelNomen;
                vm.Header.RelCalculated = data.RelCalculated;
                vm.SelectedCycleHeader = data.SelectedCycleHeader ?? string.Empty;

                vm.DataRows.Clear();
                foreach (var r in data.DataRows) vm.DataRows.Add(r);

                vm.CoordRows.Clear();
                foreach (var r in data.CoordRows) vm.CoordRows.Add(r);

                vm.Objects.Clear();
                foreach (var obj in data.Objects)
                    vm.Objects[obj.Key] = obj.Value.ToDictionary(kv => kv.Key, kv => kv.Value.ToList());

                vm.DrawingPath = dwgPath;

                _currentPath = path;
                ResetDirty();
            }
            catch (Exception ex)
            {
                _messageBox.ShowWithOptions(
                    $"Ошибка при загрузке проекта:\n{ex.Message}",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }



        private void OpenProject()
        {
            var path = _fileDialog.OpenFile("Osadka Project (*.osd)|*.osd|All Files|*.*");
            if (path == null) return;

            LoadProject(path);
        }

        private void SaveProject()
        {
            if (_currentPath == null)
            {
                SaveAsProject();
                return;
            }
            SaveTo(_currentPath);
        }

        private void SaveAsProject()
        {
            var path = _fileDialog.SaveFile("Osadka Project (*.osd)|*.osd");
            if (path == null) return;

            SaveTo(path);
            _currentPath = path;
        }

        void SaveTo(string path)
        {
            if (RawVM is not { } vm) return;

            var data = new ProjectData
            {
                Cycle = vm.Header.CycleNumber,
                MaxNomen = vm.Header.MaxNomen,
                MaxCalculated = vm.Header.MaxCalculated,
                RelNomen = vm.Header.RelNomen,
                RelCalculated = vm.Header.RelCalculated,
                SelectedCycleHeader = vm.SelectedCycleHeader,
                DataRows = vm.DataRows.ToList(),
                CoordRows = vm.CoordRows.ToList(),
                Objects = vm.Objects.ToDictionary(
                    objKv => objKv.Key,
                    objKv => objKv.Value.ToDictionary(
                        cycleKv => cycleKv.Key,
                        cycleKv => cycleKv.Value.ToList()
                    ))
            };

            _projectService.Save(path, data, vm.DrawingPath);
            ResetDirty();
        }



        private void SubscribeMeasurementRow(MeasurementRow? row)
        {
            if (row is null) return;
            if (_trackedMeasurementRows.Add(row))
            {
                row.PropertyChanged += MeasurementRow_PropertyChanged;
            }
        }

        private void UnsubscribeMeasurementRow(MeasurementRow? row)
        {
            if (row is null) return;
            if (_trackedMeasurementRows.Remove(row))
            {
                row.PropertyChanged -= MeasurementRow_PropertyChanged;
            }
        }

        private void SubscribeCoordRow(CoordRow? row)
        {
            if (row is null) return;
            if (_trackedCoordRows.Add(row))
            {
                row.PropertyChanged += CoordRow_PropertyChanged;
            }
        }

        private void UnsubscribeCoordRow(CoordRow? row)
        {
            if (row is null) return;
            if (_trackedCoordRows.Remove(row))
            {
                row.PropertyChanged -= CoordRow_PropertyChanged;
            }
        }

        private void DataRows_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Reset)
            {
                ClearMeasurementRowSubscriptions();
                foreach (var row in RawVM.DataRows)
                    SubscribeMeasurementRow(row);
            }
            else
            {
                if (e.OldItems != null)
                    foreach (MeasurementRow row in e.OldItems)
                        UnsubscribeMeasurementRow(row);

                if (e.NewItems != null)
                    foreach (MeasurementRow row in e.NewItems)
                        SubscribeMeasurementRow(row);
            }

            MarkDirty();
        }

        private void CoordRows_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Reset)
            {
                ClearCoordRowSubscriptions();
                foreach (var row in RawVM.CoordRows)
                    SubscribeCoordRow(row);
            }
            else
            {
                if (e.OldItems != null)
                    foreach (CoordRow row in e.OldItems)
                        UnsubscribeCoordRow(row);

                if (e.NewItems != null)
                    foreach (CoordRow row in e.NewItems)
                        SubscribeCoordRow(row);
            }

            MarkDirty();
        }

        private void MeasurementRow_PropertyChanged(object? sender, PropertyChangedEventArgs e)
            => MarkDirty();

        private void CoordRow_PropertyChanged(object? sender, PropertyChangedEventArgs e)
            => MarkDirty();

        private void Header_PropertyChanged(object? sender, PropertyChangedEventArgs e)
            => MarkDirty();

        private void ClearMeasurementRowSubscriptions()
        {
            foreach (var row in _trackedMeasurementRows.ToList())
                row.PropertyChanged -= MeasurementRow_PropertyChanged;
            _trackedMeasurementRows.Clear();
        }

        private void ClearCoordRowSubscriptions()
        {
            foreach (var row in _trackedCoordRows.ToList())
                row.PropertyChanged -= CoordRow_PropertyChanged;
            _trackedCoordRows.Clear();
        }

        private void MarkDirty()
        {
            if (_isDirty)
                return;

            _isDirty = true;
            if (SaveProjectCommand is RelayCommand saveRelay)
                saveRelay.NotifyCanExecuteChanged();
        }

        private void ResetDirty()
        {
            if (!_isDirty)
            {
                if (SaveProjectCommand is RelayCommand saveRelay)
                    saveRelay.NotifyCanExecuteChanged();
                return;
            }

            _isDirty = false;
            if (SaveProjectCommand is RelayCommand relay)
                relay.NotifyCanExecuteChanged();
        }

        private void DoQuickExport()
        {
            _exportService.QuickExport(
                RawVM,
                GenVM,
                RelVM,
                IncludeGeneral,
                IncludeRelative,
                IncludeGraphs);
        }
    }
}
