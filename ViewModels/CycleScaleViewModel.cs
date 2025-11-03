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
    public partial class CycleScaleViewModel : ObservableObject
    {
        private readonly RawDataViewModel _raw;
        private readonly ReadOnlyObservableCollection<CycleStateGroup> _groups;

        public ReadOnlyObservableCollection<CycleStateGroup> Groups => _groups;

        public ObservableCollection<CycleMeaningGroup> DisplayGroups { get; } = new();

        public ObservableCollection<int> CycleAxis { get; } = new();

        public IRelayCommand<CycleStateGroup> ToggleGroupCommand { get; }

        private readonly Dictionary<CycleStateKind, Brush> _brushes = new()
        {
            [CycleStateKind.Measured] = new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50)),
            [CycleStateKind.New] = new SolidColorBrush(Color.FromRgb(0x21, 0x96, 0xF3)),
            [CycleStateKind.NoAccess] = new SolidColorBrush(Color.FromRgb(0xFB, 0x8C, 0x00)),
            [CycleStateKind.Destroyed] = new SolidColorBrush(Color.FromRgb(0xE5, 0x39, 0x35)),
            [CycleStateKind.Text] = new SolidColorBrush(Color.FromRgb(0x8E, 0x24, 0xAA)),
            [CycleStateKind.Missing] = new SolidColorBrush(Color.FromRgb(0x9E, 0x9E, 0x9E))
        };

        public CycleScaleViewModel(RawDataViewModel raw)
        {
            _raw = raw ?? throw new ArgumentNullException(nameof(raw));
            _groups = new ReadOnlyObservableCollection<CycleStateGroup>(_raw.CycleGroups);
            ToggleGroupCommand = new RelayCommand<CycleStateGroup>(g => _raw.ToggleGroup(g), g => g is not null);

            _raw.CycleGroups.CollectionChanged += OnGroupsChanged;
            _raw.CycleGroupsChanged += (_, __) => OnCycleGroupsChanged();
            _raw.PropertyChanged += RawOnPropertyChanged;

            UpdateAxis();

            // Построить отрезки для уже имеющихся групп и окрасить
            foreach (var g in _raw.CycleGroups) g.RebuildSegments();
            ApplyColors();
            RefreshDisplayGroups();
        }


        private void RawOnPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(RawDataViewModel.Header) ||
                e.PropertyName == nameof(RawDataViewModel.CycleNumbers) ||
                e.PropertyName == nameof(RawDataViewModel.ObjectNumbers))
            {
                UpdateAxis();
            }
        }

        private void OnGroupsChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            ApplyColors();
            RefreshDisplayGroups();
        }

        private void OnCycleGroupsChanged()
        {
            UpdateAxis();
            foreach (var g in _raw.CycleGroups) g.RebuildSegments();
            ApplyColors();
            RefreshDisplayGroups();
            OnPropertyChanged(nameof(Groups));
        }



        private void UpdateAxis()
        {
            CycleAxis.Clear();

            // 1) Пытаемся взять явные циклы из модели, если они есть
            IEnumerable<int> cycles = Enumerable.Empty<int>();
            try
            {
                // Если у вас нет CurrentCycles — этот блок можно оставить как есть: просто перейдёт в else
                var prop = _raw.GetType().GetProperty("CurrentCycles");
                if (prop?.GetValue(_raw) is IDictionary<int, object> dict && dict.Count > 0)
                    cycles = dict.Keys;
            }
            catch
            {
                // безопасно игнорируем, используем fallback
            }

            // 2) Fallback: собираем ось по данным групп
            if (!cycles.Any())
                cycles = _raw.CycleGroups
                             .SelectMany(g => g.States)
                             .Select(s => s.CycleNumber);

            foreach (var c in cycles.Distinct().OrderBy(c => c))
                CycleAxis.Add(c);
        }

        private void ApplyColors()
        {
            foreach (var group in _raw.CycleGroups)
            {
                foreach (var state in group.States)
                    state.Brush = GetBrushForKind(state.Kind);

                foreach (var seg in group.Segments)
                    seg.Brush = GetBrushForKind(seg.Kind);
                foreach (var meaning in group.MeaningGroups)
                {
                    foreach (var seg in meaning.Segments)
                        seg.Brush = GetBrushForKind(seg.Kind);
                }
            }
        }

        private void RefreshDisplayGroups()
        {
            DisplayGroups.Clear();

            foreach (var group in _raw.CycleGroups)
            {
                foreach (var meaning in group.MeaningGroups)
                    DisplayGroups.Add(meaning);
            }
        }


        public Brush GetBrushForKind(CycleStateKind kind)
            => _brushes.TryGetValue(kind, out var brush)
                ? brush
                : Brushes.LightGray;
    }
}
