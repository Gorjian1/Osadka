using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Windows.Media;

namespace Osadka.ViewModels
{
    public partial class CycleStateViewModel : ObservableObject
    {
        private readonly RawDataViewModel _raw;
        private readonly ReadOnlyObservableCollection<CycleStateGroup> _groups;
        private readonly ObservableCollection<ColorOption> _colorOptions = new();
        private readonly HashSet<CycleStateGroup> _trackedGroups = new();
        private readonly Dictionary<string, Color> _colorCache = new(StringComparer.OrdinalIgnoreCase);
        private int _paletteIndex;

        private static readonly Color[] DefaultPalette =
        {
            Color.FromRgb(0x4A, 0x90, 0xE2),
            Color.FromRgb(0x66, 0xBB, 0x6A),
            Color.FromRgb(0xEF, 0x53, 0x50),
            Color.FromRgb(0xFF, 0xA7, 0x26),
            Color.FromRgb(0xAB, 0x47, 0xBC),
            Color.FromRgb(0x26, 0xA6, 0x9A),
            Color.FromRgb(0xEC, 0x40, 0x7A),
            Color.FromRgb(0x5C, 0x6B, 0xC0),
            Color.FromRgb(0xD4, 0xA0, 0x5B),
            Color.FromRgb(0x8D, 0x6E, 0x63)
        };

        public CycleStateViewModel(RawDataViewModel raw)
        {
            _raw = raw ?? throw new ArgumentNullException(nameof(raw));
            _groups = new ReadOnlyObservableCollection<CycleStateGroup>(_raw.CycleGroups);

            ToggleGroupCommand = new RelayCommand<CycleStateGroup>(g => _raw.ToggleGroup(g), g => g is not null);
            ChooseColorCommand = new RelayCommand<CycleStateGroup>(OnChooseColor, g => g is not null);

            _raw.CycleGroups.CollectionChanged += OnGroupsCollectionChanged;
            _raw.CycleGroupsChanged += (_, __) => RefreshAll();
            _raw.PropertyChanged += RawOnPropertyChanged;
            _raw.Header.PropertyChanged += HeaderOnPropertyChanged;

            BuildColorOptions();
            RefreshAll();
        }

        public ReadOnlyObservableCollection<CycleStateGroup> Groups => _groups;

        public ObservableCollection<CycleColumn> Columns { get; } = new();

        public ObservableCollection<CycleStateGridRow> Rows { get; } = new();

        public ObservableCollection<ColorOption> ColorOptions => _colorOptions;

        public double CellWidth => 128.0;

        public double RowHeight => 56.0;

        public IRelayCommand<CycleStateGroup> ToggleGroupCommand { get; }

        public IRelayCommand<CycleStateGroup> ChooseColorCommand { get; }

        private void BuildColorOptions()
        {
            _colorOptions.Clear();
            AddColorOption("Синий", Color.FromRgb(0x4A, 0x90, 0xE2));
            AddColorOption("Зелёный", Color.FromRgb(0x66, 0xBB, 0x6A));
            AddColorOption("Красный", Color.FromRgb(0xEF, 0x53, 0x50));
            AddColorOption("Оранжевый", Color.FromRgb(0xFF, 0xA7, 0x26));
            AddColorOption("Фиолетовый", Color.FromRgb(0xAB, 0x47, 0xBC));
            AddColorOption("Бирюзовый", Color.FromRgb(0x26, 0xA6, 0x9A));
            AddColorOption("Розовый", Color.FromRgb(0xEC, 0x40, 0x7A));
            AddColorOption("Индиго", Color.FromRgb(0x5C, 0x6B, 0xC0));
            AddColorOption("Золотой", Color.FromRgb(0xD4, 0xA0, 0x5B));
            AddColorOption("Коричневый", Color.FromRgb(0x8D, 0x6E, 0x63));
        }

        private void AddColorOption(string name, Color color)
        {
            _colorOptions.Add(new ColorOption(name, color));
        }

        private void RawOnPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(RawDataViewModel.CycleNumbers) ||
                e.PropertyName == nameof(RawDataViewModel.ObjectNumbers) ||
                e.PropertyName == nameof(RawDataViewModel.Header))
            {
                RefreshAll();
            }
        }

        private void HeaderOnPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(CycleHeader.ObjectNumber))
                RefreshAll();
        }

        private void OnGroupsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.OldItems is not null)
            {
                foreach (CycleStateGroup oldGroup in e.OldItems)
                {
                    if (_trackedGroups.Remove(oldGroup))
                        oldGroup.PropertyChanged -= GroupOnPropertyChanged;
                }
            }

            if (e.NewItems is not null)
            {
                foreach (CycleStateGroup newGroup in e.NewItems)
                {
                    if (_trackedGroups.Add(newGroup))
                        newGroup.PropertyChanged += GroupOnPropertyChanged;
                }
            }

            RefreshAll();
        }

        private void GroupOnPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (sender is not CycleStateGroup group)
                return;

            if (e.PropertyName == nameof(CycleStateGroup.DisplayColor))
            {
                _colorCache[group.Key] = group.DisplayColor;
            }
            else if (e.PropertyName == nameof(CycleStateGroup.IsEnabled))
            {
                _raw.SetGroupEnabled(group, group.IsEnabled);
            }
        }

        private void RefreshAll()
        {
            ApplyGroupColors();
            RefreshColumns();
            RefreshRows();
        }

        private void ApplyGroupColors()
        {
            foreach (var group in _raw.CycleGroups)
            {
                if (_trackedGroups.Add(group))
                    group.PropertyChanged += GroupOnPropertyChanged;

                if (!_colorCache.TryGetValue(group.Key, out var color))
                {
                    color = DefaultPalette[_paletteIndex % DefaultPalette.Length];
                    _paletteIndex++;
                    _colorCache[group.Key] = color;
                }

                if (group.DisplayColor != color)
                    group.DisplayColor = color;
                else
                    group.ApplyDisplayColor();
            }
        }

        private void RefreshColumns()
        {
            var numbers = _raw.CycleGroups
                .SelectMany(g => g.States)
                .Select(s => s.CycleNumber)
                .Distinct()
                .OrderBy(n => n)
                .ToList();

            Columns.Clear();

            foreach (var number in numbers)
            {
                _raw.CycleHeaders.TryGetValue(number, out var header);
                var caption = string.IsNullOrWhiteSpace(header) ? number.ToString() : $"{number}\n{header}";
                Columns.Add(new CycleColumn(number, caption));
            }
        }

        private void RefreshRows()
        {
            foreach (var row in Rows)
                row.Dispose();
            Rows.Clear();

            if (Columns.Count == 0)
                return;

            var cycleNumbers = Columns.Select(c => c.Number).ToList();
            foreach (var group in _raw.CycleGroups)
            {
                Rows.Add(new CycleStateGridRow(group, cycleNumbers));
            }
        }

        private void OnChooseColor(CycleStateGroup? group)
        {
            if (group is null)
                return;

            var initial = group.DisplayColor;
            using var dialog = new System.Windows.Forms.ColorDialog
            {
                AllowFullOpen = true,
                FullOpen = true,
                Color = System.Drawing.Color.FromArgb(initial.A, initial.R, initial.G, initial.B)
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var chosen = dialog.Color;
                group.DisplayColor = Color.FromArgb(chosen.A, chosen.R, chosen.G, chosen.B);
            }
        }
    }

    public sealed class CycleColumn
    {
        public CycleColumn(int number, string caption)
        {
            Number = number;
            Caption = caption;
        }

        public int Number { get; }

        public string Caption { get; }
    }

    public sealed class ColorOption
    {
        public ColorOption(string name, Color color)
        {
            Name = name;
            Color = color;
            Brush = new SolidColorBrush(color);
            if (Brush.CanFreeze)
                Brush.Freeze();
        }

        public string Name { get; }

        public Color Color { get; }

        public Brush Brush { get; }
    }

    public partial class CycleStateGridRow : ObservableObject, IDisposable
    {
        private readonly ObservableCollection<CycleStateCell> _cells = new();
        private readonly List<int> _cycles;

        public CycleStateGridRow(CycleStateGroup group, IReadOnlyList<int> cycles)
        {
            Group = group ?? throw new ArgumentNullException(nameof(group));
            _cycles = cycles?.ToList() ?? throw new ArgumentNullException(nameof(cycles));

            Cells = new ReadOnlyObservableCollection<CycleStateCell>(_cells);

            Group.PointIds.CollectionChanged += OnPointIdsChanged;
            Group.States.CollectionChanged += OnStatesChanged;
            BuildCells();
        }

        public CycleStateGroup Group { get; }

        public ReadOnlyObservableCollection<CycleStateCell> Cells { get; }

        public int PointCount => Group.PointIds.Count;

        public string PointSummary => string.Join(", ", Group.PointIds);

        private void BuildCells()
        {
            foreach (var cell in _cells)
                cell.Dispose();
            _cells.Clear();

            var map = Group.States.ToDictionary(s => s.CycleNumber);
            foreach (var cycle in _cycles)
            {
                map.TryGetValue(cycle, out var state);
                _cells.Add(new CycleStateCell(Group, cycle, state));
            }
        }

        private void OnPointIdsChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            OnPropertyChanged(nameof(PointCount));
            OnPropertyChanged(nameof(PointSummary));
        }

        private void OnStatesChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            BuildCells();
        }

        public void Dispose()
        {
            Group.PointIds.CollectionChanged -= OnPointIdsChanged;
            Group.States.CollectionChanged -= OnStatesChanged;
            foreach (var cell in _cells)
                cell.Dispose();
            _cells.Clear();
        }
    }

    public sealed class CycleStateCell : ObservableObject, IDisposable
    {
        public CycleStateCell(CycleStateGroup group, int cycleNumber, CycleState? state)
        {
            if (group is null)
                throw new ArgumentNullException(nameof(group));

            CycleNumber = cycleNumber;
            State = state;

            if (State is not null)
                State.PropertyChanged += OnStatePropertyChanged;
        }

        public int CycleNumber { get; }

        public CycleState? State { get; }

        public bool HasData => State is { Kind: not CycleStateKind.Missing };

        public string Label
        {
            get
            {
                if (State is null)
                    return string.Empty;

                if (!string.IsNullOrWhiteSpace(State.Annotation))
                    return State.Annotation!;

                return State.Kind switch
                {
                    CycleStateKind.Measured => "Измерено",
                    CycleStateKind.New => "Новая",
                    CycleStateKind.NoAccess => "Нет доступа",
                    CycleStateKind.Destroyed => "Уничтожена",
                    CycleStateKind.Text => "Примечание",
                    CycleStateKind.Missing => "Нет данных",
                    _ => State.Kind.ToString()
                };
            }
        }

        public Brush Fill => State?.Brush ?? Brushes.Transparent;

        private void OnStatePropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(CycleState.Brush))
                OnPropertyChanged(nameof(Fill));
        }

        public void Dispose()
        {
            if (State is not null)
                State.PropertyChanged -= OnStatePropertyChanged;
        }
    }
}
