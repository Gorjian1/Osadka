using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using CommunityToolkit.Mvvm.Messaging;
using Microsoft.VisualBasic;
using Osadka.Core.Units;
using Osadka.Messages;
using Osadka.Models;
using Osadka.Services;
using Osadka.Services.Abstractions;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
            _settings.TemplatePath = value;
            _settings.Save();
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

        // Преобразование старых значений в новые через отношение масштабов
        partial void OnCoordUnitChanged(CoordUnits oldVal, CoordUnits newVal)
        {
            var oldU = Map(oldVal);
            var newU = Map(newVal);
            double k = UnitConverter.ToMm(1.0, newU) / UnitConverter.ToMm(1.0, oldU);

            foreach (var p in CoordRows)
            {
                p.X *= k;
                p.Y *= k;
            }
               foreach (var r in DataRows)
                   {
                       if (r.Mark is double m) r.Mark = m * k;
                       if (r.Settl is double s) r.Settl = s * k;
                       if (r.Total is double t) r.Total = t * k;
                   }
            // Держим SourceUnit согласованным с CoordUnit
            SourceUnit = newU;
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
        private readonly Osadka.Services.Abstractions.ISettingsService _settings;
        private readonly Osadka.Services.Abstractions.IExcelImportService _excelImport;
        private readonly Osadka.Services.Abstractions.IMessageBoxService _messageBox;
        private readonly Osadka.Services.Abstractions.IFileDialogService _fileDialog;

        public IRelayCommand OpenTemplate { get; }
        public IRelayCommand ChooseOrOpenTemplateCommand { get; }
        public IRelayCommand ClearTemplateCommand { get; }
        public IRelayCommand LoadFromWorkbookCommand { get; }
        public IRelayCommand ClearCommand { get; }

        public RawDataViewModel(
            Osadka.Services.Abstractions.ISettingsService settings,
            Osadka.Services.Abstractions.IExcelImportService excelImport,
            Osadka.Services.Abstractions.IMessageBoxService messageBox,
            Osadka.Services.Abstractions.IFileDialogService fileDialog)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _excelImport = excelImport ?? throw new ArgumentNullException(nameof(excelImport));
            _messageBox = messageBox ?? throw new ArgumentNullException(nameof(messageBox));
            _fileDialog = fileDialog ?? throw new ArgumentNullException(nameof(fileDialog));

            OpenTemplate = new RelayCommand(OpenTemplatePicker);
            ChooseOrOpenTemplateCommand = new RelayCommand(ChooseOrOpenTemplate);
            ClearTemplateCommand = new RelayCommand(ClearTemplate, () => HasCustomTemplate);

            LoadFromWorkbookCommand = new RelayCommand(OnLoadWorkbook);
            ClearCommand = new RelayCommand(OnClear);

            _settings.Load();
            TemplatePath = _settings.TemplatePath;

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
                    _messageBox.ShowWithOptions(
                        $"Не удалось открыть файл шаблона:\n{ex.Message}",
                        "Шаблон",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
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
            _messageBox.ShowWithOptions(
                "Путь к пользовательскому шаблону очищен. Будет использован встроенный template.xlsx.",
                "Шаблон",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        // Диалог выбора шаблона
        private void OpenTemplatePicker()
        {
            var filePath = _fileDialog.OpenFile(
                "Excel шаблоны (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Все файлы|*.*");

            if (filePath != null)
            {
                TemplatePath = filePath;
                _messageBox.ShowWithOptions(
                    "Шаблон успешно выбран:\n" + TemplatePath,
                    "Шаблон",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
        }

        // === Импорт из Excel ===
        private void OnLoadWorkbook()
        {
            var filePath = _fileDialog.OpenFile("Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm");

            if (filePath != null)
                ImportFromExcel(filePath);
        }

        public void ImportFromExcel(string filePath)
        {
            try
            {
                var result = _excelImport.ImportFromExcel(filePath, Map(CoordUnit));
                if (result == null)
                    return; // Пользователь отменил

                // Применяем результаты импорта
                _objects.Clear();
                foreach (var obj in result.Objects)
                    _objects[obj.Key] = obj.Value;

                _cycleHeaders.Clear();
                foreach (var hdr in result.CycleHeaders)
                    _cycleHeaders[hdr.Key] = hdr.Value;

                // Обновляем списки объектов
                ObjectNumbers.Clear();
                foreach (var k in _objects.Keys.OrderBy(k => k))
                    ObjectNumbers.Add(k);

                // Устанавливаем рекомендуемый объект
                Header.ObjectNumber = ObjectNumbers.Contains(result.SuggestedObjectNumber)
                    ? result.SuggestedObjectNumber
                    : (ObjectNumbers.Count > 0 ? ObjectNumbers[0] : 1);

                // Обновляем списки циклов
                CycleNumbers.Clear();
                if (_objects.TryGetValue(Header.ObjectNumber, out var cyclesForObject))
                {
                    foreach (var k in cyclesForObject.Keys.OrderBy(k => k))
                        CycleNumbers.Add(k);
                }

                // Устанавливаем рекомендуемый цикл
                if (CycleNumbers.Count > 0)
                {
                    Header.CycleNumber = CycleNumbers.Contains(result.SuggestedCycleNumber)
                        ? result.SuggestedCycleNumber
                        : CycleNumbers[0];
                }

                RefreshData();
            }
            catch (Exception ex)
            {
                _messageBox.ShowWithOptions(
                    $"Ошибка при импорте Excel:\n{ex.Message}",
                    "Импорт",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
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
