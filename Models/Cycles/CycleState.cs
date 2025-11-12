using CommunityToolkit.Mvvm.ComponentModel;
using System.Windows.Media;

namespace Osadka.Models.Cycles;

public partial class CycleState : ObservableObject
{
    public CycleState(int cycleNumber, CycleStateKind kind, string? annotation)
    {
        CycleNumber = cycleNumber;
        Kind = kind;
        Annotation = annotation;
    }

    public int CycleNumber { get; }
    public CycleStateKind Kind { get; }
    public string? Annotation { get; }
    public bool HasData => Kind != CycleStateKind.Missing;

    [ObservableProperty]
    private Brush _brush = Brushes.Transparent;
}
