using CommunityToolkit.Mvvm.ComponentModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;

namespace Osadka.Models.Cycles;

public partial class CycleMeaningGroup : ObservableObject, IDisposable
{
    public CycleMeaningGroup(CycleStateGroup owner, CycleStateKind kind, IEnumerable<CycleSegment> segments)
    {
        Owner = owner ?? throw new ArgumentNullException(nameof(owner));
        Kind = kind;
        Segments = new ObservableCollection<CycleSegment>(segments ?? Enumerable.Empty<CycleSegment>());
        Title = ComposeTitle();
        Owner.PropertyChanged += OwnerOnPropertyChanged;
    }

    public CycleStateGroup Owner { get; }

    public CycleStateKind Kind { get; }

    public ObservableCollection<CycleSegment> Segments { get; }

    [ObservableProperty]
    private string _title = string.Empty;

    [ObservableProperty]
    private int _displayOrder;

    public bool IsAlternateRow => DisplayOrder % 2 == 1;

    public bool IsFirstRow => DisplayOrder == 0;

    public string Meaning => Kind switch
    {
        CycleStateKind.Measured => "Измерено",
        CycleStateKind.New => "Новая точка",
        CycleStateKind.NoAccess => "Нет доступа",
        CycleStateKind.Destroyed => "Уничтожена",
        CycleStateKind.Text => "Особая отметка",
        CycleStateKind.Missing => "Нет данных",
        _ => Kind.ToString()
    };

    private string ComposeTitle()
    {
        string prefix = string.IsNullOrWhiteSpace(Owner.DisplayName)
            ? string.Empty
            : Owner.DisplayName + " — ";

        return prefix + Meaning;
    }

    private void OwnerOnPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(CycleStateGroup.DisplayName))
            Title = ComposeTitle();
    }

    partial void OnDisplayOrderChanged(int value)
    {
        OnPropertyChanged(nameof(IsAlternateRow));
        OnPropertyChanged(nameof(IsFirstRow));
    }

    public void Dispose()
    {
        Owner.PropertyChanged -= OwnerOnPropertyChanged;
    }
}
