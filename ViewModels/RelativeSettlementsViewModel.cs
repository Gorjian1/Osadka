using System;
using System.Linq;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ClosedXML.Excel;
using Microsoft.Win32;
using Osadka.Models;
using Osadka.Services;

namespace Osadka.ViewModels
{
    public partial class RelativeSettlementsViewModel : ObservableObject
    {
        public ObservableCollection<RelativeRow> AllRows { get; } = new();
        public ObservableCollection<RelativeRow> ExceededSpRows { get; } = new();
        public ObservableCollection<RelativeRow> ExceededCalcRows { get; } = new();

        public IRelayCommand RecalcCommand { get; }


        private readonly RawDataViewModel _raw;
        private readonly RelativeReportService _relSvc;

        public RelativeSettlementsViewModel(RawDataViewModel raw, RelativeReportService svc)
        {
            _raw = raw ?? throw new ArgumentNullException(nameof(raw));
            _relSvc = svc ?? throw new ArgumentNullException(nameof(svc));

            RecalcCommand = new RelayCommand(Recalc);

            _raw.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName is nameof(RawDataViewModel.Header.ObjectNumber)
                                     or nameof(RawDataViewModel.DataRows)
                                     or nameof(RawDataViewModel.CoordRows))
                    Recalc();
            };
            _raw.DataRows.CollectionChanged += (_, __) => Recalc();
            _raw.CoordRows.CollectionChanged += (_, __) => Recalc();

            Recalc();
        }

        private void Recalc()
        {
            double spLim = _raw.Header.RelNomen ?? 0;
            double calcLim = _raw.Header.RelCalculated ?? 0;

            var report = _relSvc.Build(
                _raw.CoordRows,
                _raw.DataRows,
                spLim,
                calcLim);

            AllRows.Reset(report.AllRows);
            ExceededSpRows.Reset(report.ExceededSpRows);
            ExceededCalcRows.Reset(report.ExceededCalcRows);

        }

    }

    static class Extensions
    {
        public static void Reset<T>(this ObservableCollection<T> coll, IEnumerable<T> src)
        {
            coll.Clear();
            foreach (var i in src) coll.Add(i);
        }
    }
}
