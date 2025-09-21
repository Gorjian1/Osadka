using System;
using System.Linq;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
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

        // Отчёт из бизнес-логики Relative (для экспорта/тегов)
        public RelativeReport? Report { get; private set; }
        public double? RelativeMaxValue =>
            (Report?.MaxRelative is { } e && !double.IsNaN(e.Value)) ? e.Value : (double?)null;
        public IReadOnlyList<string> RelativeMaxIdPairs => Report?.MaxRelative?.Ids ?? Array.Empty<string>();

        public RelativeSettlementsViewModel(RawDataViewModel raw, RelativeReportService svc)
        {
            _raw = raw ?? throw new ArgumentNullException(nameof(raw));
            _relSvc = svc ?? throw new ArgumentNullException(nameof(svc));

            RecalcCommand = new RelayCommand(Recalc);

            // Пересчёт при изменениях входных данных
            _raw.PropertyChanged += RawOnPropertyChanged;
            _raw.DataRows.CollectionChanged += (_, __) => Recalc();
            _raw.CoordRows.CollectionChanged += (_, __) => Recalc();

            Recalc();
        }

        private void RawOnPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName is nameof(RawDataViewModel.Header)
                || e.PropertyName is nameof(RawDataViewModel.DataRows)
                || e.PropertyName is nameof(RawDataViewModel.CoordRows))
            {
                Recalc();
            }
        }

        private static bool HasNumericTotal(MeasurementRow r)
            => r.Total.HasValue && !double.IsNaN(r.Total.Value);

        private void Recalc()
        {
            // Пороговые значения
            double spLim = _raw.Header.RelNomen ?? 0;
            double calcLim = _raw.Header.RelCalculated ?? 0;

            // 1) Берём срез данных (включая прочерки)
            var data = _raw.DataRows.ToList();
            var coords = _raw.CoordRows.ToList();

            // 2) ВЫРАВНИВАНИЕ ТОЛЬКО ПО ПОЗИЦИИ:
            //    координаты "прочёркнутых" точек НЕ удаляем, берём по индексу
            var coordsAligned = new List<CoordRow>(data.Count);
            for (int i = 0; i < data.Count; i++)
            {
                if (i < coords.Count)
                    coordsAligned.Add(coords[i]);
                else
                    coordsAligned.Add(new CoordRow { X = double.NaN, Y = double.NaN, Id = data[i].Id });
            }

            // 3) Расчёт (без фильтраций — пары строго по индексам)
            var report = _relSvc.Build(coordsAligned, data, spLim, calcLim);

            // 4) СОХРАНИТЬ отчёт во VM (важно для экспорта/тегов)
            Report = report;
            OnPropertyChanged(nameof(Report));
            OnPropertyChanged(nameof(RelativeMaxValue));
            OnPropertyChanged(nameof(RelativeMaxIdPairs));

            // 5) Обновить таблицы
            AllRows.Reset(report.AllRows);
            ExceededSpRows.Reset(report.ExceededSpRows);
            ExceededCalcRows.Reset(report.ExceededCalcRows);
        }


    }

    internal static class CollectionExtensions
    {
        public static void Reset<T>(this ObservableCollection<T> coll, IEnumerable<T> src)
        {
            coll.Clear();
            foreach (var i in src) coll.Add(i);
        }
    }
}
