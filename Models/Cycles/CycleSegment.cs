using CommunityToolkit.Mvvm.ComponentModel;
using System.Windows.Media;

namespace Osadka.Models.Cycles;

public partial class CycleSegment : ObservableObject
{
    public CycleSegment(int startIndex, int endIndex, int cycleFrom, int cycleTo, CycleStateKind kind, string? annotation)
    {
        StartIndex = startIndex;
        Span = endIndex - startIndex + 1;
        CycleFrom = cycleFrom;
        CycleTo = cycleTo;
        Kind = kind;
        Annotation = annotation;
    }

    public int StartIndex { get; }
    public int Span { get; }
    public int CycleFrom { get; }
    public int CycleTo { get; }
    public CycleStateKind Kind { get; }
    public string? Annotation { get; }

    [ObservableProperty] private Brush _brush = Brushes.Transparent;
}
