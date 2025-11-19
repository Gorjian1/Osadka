using CommunityToolkit.Mvvm.ComponentModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Media;

namespace Osadka.Models.Cycles;

public partial class CycleStateGroup : ObservableObject
{
    public CycleStateGroup(string key, IEnumerable<CycleState> states)
    {
        if (string.IsNullOrEmpty(key))
            throw new ArgumentException("Key cannot be null or empty", nameof(key));

        Key = key;
        States = new ObservableCollection<CycleState>(states ?? Enumerable.Empty<CycleState>());
    }

    public string Key { get; }

    [ObservableProperty]
    private string _displayName = string.Empty;

    public ObservableCollection<string> PointIds { get; } = new();

    public ObservableCollection<CycleState> States { get; }

    // Отрезки для Ганта (склеенные одинаковые статусы без Missing)
    public ObservableCollection<CycleSegment> Segments { get; } = new();

    // Разделение отрезков по смыслу (типу статуса)
    public ObservableCollection<CycleMeaningGroup> MeaningGroups { get; } = new();

    [ObservableProperty]
    private bool _isEnabled = true;

    private bool _suppressColorFlag;

    [ObservableProperty]
    private Color _accentColor = Color.FromRgb(0x4C, 0xAF, 0x50);

    [ObservableProperty]
    private bool _hasCustomColor;

    public SolidColorBrush AccentBrush => new SolidColorBrush(AccentColor);

    public void SetAccentColor(Color color, bool markCustom)
    {
        _suppressColorFlag = true;
        AccentColor = color;
        HasCustomColor = markCustom;
        _suppressColorFlag = false;
        OnPropertyChanged(nameof(AccentBrush));
    }

    partial void OnAccentColorChanged(Color value)
    {
        if (!_suppressColorFlag)
        {
            HasCustomColor = true;
        }

        OnPropertyChanged(nameof(AccentBrush));
    }

    public void RebuildSegments()
    {
        foreach (var meaning in MeaningGroups)
            meaning.Dispose();

        MeaningGroups.Clear();
        Segments.Clear();
        if (States.Count == 0)
            return;

        int start = 0;
        var kind = States[0].Kind;
        string? ann = States[0].Annotation;

        var segments = new List<CycleSegment>();

        for (int i = 1; i < States.Count; i++)
        {
            var s = States[i];
            if (s.Kind == kind) continue;

            if (kind != CycleStateKind.Missing)
                segments.Add(new CycleSegment(
                    start, i - 1,
                    States[start].CycleNumber,
                    States[i - 1].CycleNumber,
                    kind, ann));

            start = i;
            kind = s.Kind;
            ann = s.Annotation;
        }

        if (kind != CycleStateKind.Missing)
            segments.Add(new CycleSegment(
                start, States.Count - 1,
                States[start].CycleNumber,
                States[States.Count - 1].CycleNumber,
                kind, ann));

        foreach (var segment in segments)
            Segments.Add(segment);

        if (segments.Count == 0)
            return;

        var order = new List<CycleStateKind>();
        foreach (var segment in segments)
        {
            if (!order.Contains(segment.Kind))
                order.Add(segment.Kind);
        }

        foreach (var k in order)
        {
            var sameKind = segments.Where(s => s.Kind == k).ToList();
            if (sameKind.Count == 0)
                continue;

            MeaningGroups.Add(new CycleMeaningGroup(this, k, sameKind));
        }
    }

    public void SortPointIds(IComparer<string> comparer)
    {
        if (comparer is null)
            throw new ArgumentNullException(nameof(comparer));

        if (PointIds.Count <= 1)
            return;

        var ordered = PointIds.OrderBy(id => id, comparer).ToList();

        if (ordered.SequenceEqual(PointIds))
            return;

        PointIds.Clear();
        foreach (var id in ordered)
            PointIds.Add(id);
    }
}
