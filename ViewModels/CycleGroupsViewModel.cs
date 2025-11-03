using CommunityToolkit.Mvvm.ComponentModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Media;

namespace Osadka.ViewModels
{
    public partial class CycleGroupsViewModel : ObservableObject, IDisposable
    {
        private static readonly Color[] DefaultPalette =
        {
            Color.FromRgb(0x3F, 0x51, 0xB5),
            Color.FromRgb(0x00, 0x96, 0x88),
            Color.FromRgb(0x5E, 0x35, 0xB1),
            Color.FromRgb(0x03, 0xA9, 0xF4),
            Color.FromRgb(0x8B, 0xC3, 0x4A),
            Color.FromRgb(0xFF, 0xB3, 0x3A),
            Color.FromRgb(0xE5, 0x39, 0x35),
            Color.FromRgb(0x46, 0x82, 0xB4)
        };

        private readonly RawDataViewModel _raw;

        public ObservableCollection<CycleColumnHeader> Columns { get; } = new();
        public ObservableCollection<CycleGroupRow> Rows { get; } = new();
        public ObservableCollection<CycleGroupSettingsEntry> Settings { get; } = new();
        public ObservableCollection<ColorOption> ColorOptions { get; }

        public bool HasData => Rows.Count > 0;
        public int CycleCount => Columns.Count;

        public CycleGroupsViewModel(RawDataViewModel raw)
        {
            _raw = raw ?? throw new ArgumentNullException(nameof(raw));

            ColorOptions = new ObservableCollection<ColorOption>(BuildColorOptions());

            Columns.CollectionChanged += ColumnsOnCollectionChanged;

            _raw.CycleGroups.CollectionChanged += OnGroupsChanged;
            _raw.CycleGroupsChanged += OnCycleGroupsChanged;
            _raw.PropertyChanged += RawOnPropertyChanged;

            Refresh();
        }

        private static IEnumerable<ColorOption> BuildColorOptions()
        {
            var baseOptions = new List<ColorOption>
            {
                new("Синий", Color.FromRgb(0x3F, 0x51, 0xB5)),
                new("Бирюзовый", Color.FromRgb(0x00, 0x96, 0x88)),
                new("Фиолетовый", Color.FromRgb(0x5E, 0x35, 0xB1)),
                new("Голубой", Color.FromRgb(0x03, 0xA9, 0xF4)),
                new("Зелёный", Color.FromRgb(0x4C, 0xAF, 0x50)),
                new("Лайм", Color.FromRgb(0x8B, 0xC3, 0x4A)),
                new("Жёлтый", Color.FromRgb(0xFF, 0xB3, 0x3A)),
                new("Оранжевый", Color.FromRgb(0xFB, 0x8C, 0x00)),
                new("Красный", Color.FromRgb(0xE5, 0x39, 0x35)),
                new("Серый", Color.FromRgb(0x90, 0xA4, 0xAE))
            };

            return baseOptions.DistinctBy(o => o.Color).ToList();
        }

        private void RawOnPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(RawDataViewModel.CycleNumbers) ||
                e.PropertyName == nameof(RawDataViewModel.Header))
            {
                Refresh();
            }
        }

        private void OnCycleGroupsChanged(object? sender, EventArgs e) => Refresh();

        private void OnGroupsChanged(object? sender, NotifyCollectionChangedEventArgs e) => Refresh();

        private void Refresh()
        {
            var cycles = _raw.CycleNumbers.Any()
                ? _raw.CycleNumbers.OrderBy(c => c).ToList()
                : _raw.CycleGroups
                      .SelectMany(g => g.States)
                      .Select(s => s.CycleNumber)
                      .Distinct()
                      .OrderBy(c => c)
                      .ToList();

            foreach (var row in Rows)
                row.Dispose();
            Rows.Clear();

            foreach (var setting in Settings)
                setting.Dispose();
            Settings.Clear();

            UpdateColumns(cycles);
            AssignDefaultColors();

            foreach (var group in _raw.CycleGroups)
            {
                var row = new CycleGroupRow(group, cycles);
                Rows.Add(row);
                Settings.Add(new CycleGroupSettingsEntry(group, _raw));
            }

            OnPropertyChanged(nameof(HasData));
        }

        private void UpdateColumns(IReadOnlyList<int> cycles)
        {
            Columns.Clear();
            foreach (var cycle in cycles)
            {
                string label = _raw.CycleHeaders.TryGetValue(cycle, out var text) && !string.IsNullOrWhiteSpace(text)
                    ? text
                    : $"Цикл {cycle}";
                Columns.Add(new CycleColumnHeader(cycle, label));
            }

            OnPropertyChanged(nameof(CycleCount));
        }

        private void AssignDefaultColors()
        {
            if (_raw.CycleGroups.Count == 0)
                return;

            int index = 0;
            foreach (var group in _raw.CycleGroups)
            {
                if (group.HasCustomColor)
                    continue;

                var color = DefaultPalette[index++ % DefaultPalette.Length];
                group.SetAccentColor(color, false);
            }
        }

        public void Dispose()
        {
            Columns.CollectionChanged -= ColumnsOnCollectionChanged;
            _raw.CycleGroups.CollectionChanged -= OnGroupsChanged;
            _raw.CycleGroupsChanged -= OnCycleGroupsChanged;
            _raw.PropertyChanged -= RawOnPropertyChanged;

            foreach (var row in Rows)
                row.Dispose();
            Rows.Clear();

            foreach (var setting in Settings)
                setting.Dispose();
            Settings.Clear();
        }

        private void ColumnsOnCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            OnPropertyChanged(nameof(CycleCount));
        }
    }

    public sealed class ColorOption
    {
        public ColorOption(string name, Color color)
        {
            Name = name;
            Color = color;
            Brush = new SolidColorBrush(color);
        }

        public string Name { get; }
        public Color Color { get; }
        public Brush Brush { get; }
    }

    public sealed class CycleColumnHeader
    {
        public CycleColumnHeader(int number, string label)
        {
            Number = number;
            Label = label;
        }

        public int Number { get; }
        public string Label { get; }
    }

    public partial class CycleGroupSettingsEntry : ObservableObject, IDisposable
    {
        private readonly RawDataViewModel _raw;

        public CycleStateGroup Group { get; }

        public CycleGroupSettingsEntry(CycleStateGroup group, RawDataViewModel raw)
        {
            Group = group ?? throw new ArgumentNullException(nameof(group));
            _raw = raw ?? throw new ArgumentNullException(nameof(raw));

            Group.PropertyChanged += GroupOnPropertyChanged;
            Group.PointIds.CollectionChanged += PointIdsOnCollectionChanged;
        }

        public string Name => Group.DisplayName;
        public int PointCount => Group.PointIds.Count;
        public IEnumerable<string> PointIds => Group.PointIds;

        public bool IsIncluded
        {
            get => Group.IsEnabled;
            set
            {
                if (value == Group.IsEnabled)
                    return;

                _raw.SetGroupEnabled(Group, value);
                OnPropertyChanged();
            }
        }

        public Color AccentColor
        {
            get => Group.AccentColor;
            set
            {
                if (value == Group.AccentColor)
                    return;

                Group.SetAccentColor(value, true);
                OnPropertyChanged();
                OnPropertyChanged(nameof(AccentBrush));
            }
        }

        public Brush AccentBrush => Group.AccentBrush;

        public string PointSummary
        {
            get
            {
                if (Group.PointIds.Count == 0)
                    return "Нет точек";

                const int PreviewLimit = 4;
                var preview = Group.PointIds.Take(PreviewLimit).ToList();
                string joined = string.Join(", ", preview);
                if (Group.PointIds.Count > PreviewLimit)
                    joined += $", … (ещё {Group.PointIds.Count - PreviewLimit})";
                return joined;
            }
        }

        private void GroupOnPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(CycleStateGroup.DisplayName))
            {
                OnPropertyChanged(nameof(Name));
            }
            else if (e.PropertyName == nameof(CycleStateGroup.IsEnabled))
            {
                OnPropertyChanged(nameof(IsIncluded));
            }
            else if (e.PropertyName == nameof(CycleStateGroup.AccentColor))
            {
                OnPropertyChanged(nameof(AccentColor));
                OnPropertyChanged(nameof(AccentBrush));
            }
        }

        private void PointIdsOnCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            OnPropertyChanged(nameof(PointCount));
            OnPropertyChanged(nameof(PointSummary));
        }

        public void Dispose()
        {
            Group.PropertyChanged -= GroupOnPropertyChanged;
            Group.PointIds.CollectionChanged -= PointIdsOnCollectionChanged;
        }
    }

    public partial class CycleGroupRow : ObservableObject, IDisposable
    {
        private readonly IReadOnlyList<int> _cycles;

        public CycleStateGroup Group { get; }
        public ObservableCollection<CycleDisplayCell> Cells { get; } = new();
        public ObservableCollection<CycleGroupSegment> Segments { get; } = new();

        public CycleGroupRow(CycleStateGroup group, IReadOnlyList<int> cycles)
        {
            Group = group ?? throw new ArgumentNullException(nameof(group));
            _cycles = cycles ?? throw new ArgumentNullException(nameof(cycles));

            BuildCells();

            Group.PropertyChanged += GroupOnPropertyChanged;
            Group.PointIds.CollectionChanged += PointIdsOnCollectionChanged;
            Group.States.CollectionChanged += StatesOnCollectionChanged;
        }

        public string Name => Group.DisplayName;
        public int PointCount => Group.PointIds.Count;
        public bool IsIncluded => Group.IsEnabled;
        public Brush Accent => Group.AccentBrush;

        private void BuildCells()
        {
            Cells.Clear();
            var map = Group.States.ToDictionary(s => s.CycleNumber, s => s);
            foreach (var cycle in _cycles)
            {
                map.TryGetValue(cycle, out var state);
                Cells.Add(new CycleDisplayCell(Group, state, cycle));
            }

            BuildSegments();
        }

        private void BuildSegments()
        {
            Segments.Clear();

            if (Cells.Count == 0)
                return;

            int segmentStart = 0;
            CycleStateKind currentKind = Cells[0].Kind;

            for (int index = 1; index <= Cells.Count; index++)
            {
                bool isBoundary = index == Cells.Count || Cells[index].Kind != currentKind;
                if (!isBoundary)
                    continue;

                AddSegment(segmentStart, index - 1, currentKind);

                if (index < Cells.Count)
                {
                    segmentStart = index;
                    currentKind = Cells[index].Kind;
                }
            }
        }

        private void AddSegment(int startIndex, int endIndex, CycleStateKind kind)
        {
            var firstCell = Cells[startIndex];
            var lastCell = Cells[endIndex];
            int length = endIndex - startIndex + 1;

            var annotations = Cells
                .Skip(startIndex)
                .Take(length)
                .Select(c => c.Annotation)
                .Where(a => !string.IsNullOrWhiteSpace(a))
                .ToList();

            string label = CycleDisplayFormatting.GetLabel(kind);
            string toolTip = CycleDisplayFormatting.GetSegmentToolTip(
                firstCell.CycleNumber,
                lastCell.CycleNumber,
                label,
                annotations);

            Segments.Add(new CycleGroupSegment(
                startIndex,
                length,
                firstCell.CycleNumber,
                lastCell.CycleNumber,
                kind,
                firstCell.Background,
                firstCell.Foreground,
                label,
                toolTip));
        }

        private void GroupOnPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(CycleStateGroup.DisplayName))
            {
                OnPropertyChanged(nameof(Name));
            }
            else if (e.PropertyName == nameof(CycleStateGroup.IsEnabled))
            {
                OnPropertyChanged(nameof(IsIncluded));
            }
            else if (e.PropertyName == nameof(CycleStateGroup.AccentColor))
            {
                foreach (var cell in Cells)
                    cell.RefreshAppearance();
                OnPropertyChanged(nameof(Accent));
                BuildSegments();
            }
        }

        private void PointIdsOnCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            OnPropertyChanged(nameof(PointCount));
        }

        private void StatesOnCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            BuildCells();
        }

        public void Dispose()
        {
            Group.PropertyChanged -= GroupOnPropertyChanged;
            Group.PointIds.CollectionChanged -= PointIdsOnCollectionChanged;
            Group.States.CollectionChanged -= StatesOnCollectionChanged;
        }
    }

    public sealed class CycleGroupSegment
    {
        public CycleGroupSegment(
            int startIndex,
            int length,
            int startCycle,
            int endCycle,
            CycleStateKind kind,
            Brush background,
            Brush foreground,
            string label,
            string toolTip)
        {
            StartIndex = startIndex;
            Length = length;
            StartCycle = startCycle;
            EndCycle = endCycle;
            Kind = kind;
            Background = background;
            Foreground = foreground;
            Label = label;
            ToolTip = toolTip;
        }

        public int StartIndex { get; }
        public int Length { get; }
        public int StartCycle { get; }
        public int EndCycle { get; }
        public CycleStateKind Kind { get; }
        public Brush Background { get; }
        public Brush Foreground { get; }
        public string Label { get; }
        public string ToolTip { get; }
    }

    public partial class CycleDisplayCell : ObservableObject
    {
        private readonly CycleStateGroup _group;
        private readonly CycleState? _state;

        public CycleDisplayCell(CycleStateGroup group, CycleState? state, int cycleNumber)
        {
            _group = group ?? throw new ArgumentNullException(nameof(group));
            _state = state;
            CycleNumber = cycleNumber;

            RefreshAppearance();
        }

        public int CycleNumber { get; }
        public CycleStateKind Kind => _state?.Kind ?? CycleStateKind.Missing;
        public string? Annotation => _state?.Annotation;

        [ObservableProperty]
        private Brush _background = Brushes.Transparent;

        [ObservableProperty]
        private Brush _foreground = Brushes.Gray;

        [ObservableProperty]
        private string _label = "—";

        [ObservableProperty]
        private string _toolTip = string.Empty;

        public void RefreshAppearance()
        {
            if (_state is null || _state.Kind == CycleStateKind.Missing)
            {
                Label = "—";
                Background = Brushes.Transparent;
                Foreground = Brushes.Gray;
                ToolTip = $"Цикл {CycleNumber}: нет данных";
                return;
            }

            Label = CycleDisplayFormatting.GetLabel(_state.Kind);
            var color = CycleDisplayFormatting.GetFill(_group.AccentColor, _state.Kind);
            Background = new SolidColorBrush(color);
            Foreground = CycleDisplayFormatting.GetForeground(color);
            ToolTip = CycleDisplayFormatting.GetToolTip(CycleNumber, Label, _state.Annotation);
        }
    }

    internal static class CycleDisplayFormatting
    {
        public static string GetLabel(CycleStateKind kind) => kind switch
        {
            CycleStateKind.Measured => "Изм",
            CycleStateKind.New => "Нов",
            CycleStateKind.NoAccess => "Нет",
            CycleStateKind.Destroyed => "Разр",
            CycleStateKind.Text => "Прим",
            _ => "—"
        };

        public static string GetToolTip(int cycleNumber, string label, string? annotation)
        {
            if (string.IsNullOrWhiteSpace(annotation))
                return $"Цикл {cycleNumber}: {label}";

            return $"Цикл {cycleNumber}: {label}\n{annotation}";
        }

        public static string GetSegmentToolTip(int startCycle, int endCycle, string label, IEnumerable<string?> annotations)
        {
            var builder = new StringBuilder();
            if (startCycle == endCycle)
            {
                builder.Append($"Цикл {startCycle}: {label}");
            }
            else
            {
                builder.Append($"Циклы {startCycle}–{endCycle}: {label}");
            }

            var notes = annotations
                .Where(a => !string.IsNullOrWhiteSpace(a))
                .Select(a => a!.Trim())
                .ToList();

            if (notes.Count > 0)
            {
                builder.AppendLine();
                builder.Append(string.Join(Environment.NewLine, notes));
            }

            return builder.ToString();
        }

        public static Color GetFill(Color baseColor, CycleStateKind kind)
        {
            if (kind == CycleStateKind.Missing)
                return Colors.Transparent;

            if (baseColor.A == 0)
                baseColor = Color.FromRgb(0x4C, 0xAF, 0x50);

            return kind switch
            {
                CycleStateKind.Measured => baseColor,
                CycleStateKind.New => Blend(baseColor, Colors.White, 0.25),
                CycleStateKind.NoAccess => Blend(baseColor, Colors.Gold, 0.35),
                CycleStateKind.Destroyed => Blend(baseColor, Colors.Black, 0.35),
                CycleStateKind.Text => Blend(baseColor, Colors.White, 0.45),
                _ => baseColor
            };
        }

        public static Brush GetForeground(Color color)
        {
            double luminance = (0.299 * color.R + 0.587 * color.G + 0.114 * color.B) / 255.0;
            return luminance > 0.55 ? Brushes.Black : Brushes.White;
        }

        private static Color Blend(Color source, Color target, double amount)
        {
            amount = Math.Clamp(amount, 0, 1);
            byte r = (byte)(source.R + (target.R - source.R) * amount);
            byte g = (byte)(source.G + (target.G - source.G) * amount);
            byte b = (byte)(source.B + (target.B - source.B) * amount);
            return Color.FromRgb(r, g, b);
        }
    }
}
