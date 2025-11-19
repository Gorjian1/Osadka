using Osadka.ViewModels;

namespace Osadka.Services.Abstractions;

/// <summary>
/// Сервис для экспорта данных в Excel
/// </summary>
public interface IExportService
{
    /// <summary>
    /// Выполняет быстрый экспорт отчётов в Excel
    /// </summary>
    /// <param name="rawDataViewModel">Данные измерений</param>
    /// <param name="generalReportViewModel">Общий отчёт</param>
    /// <param name="relativeSettlementsViewModel">Относительные осадки</param>
    /// <param name="includeGeneral">Включить общий отчёт</param>
    /// <param name="includeRelative">Включить относительную разность</param>
    /// <param name="includeGraphs">Включить графики динамики</param>
    void QuickExport(
        RawDataViewModel rawDataViewModel,
        GeneralReportViewModel generalReportViewModel,
        RelativeSettlementsViewModel relativeSettlementsViewModel,
        bool includeGeneral,
        bool includeRelative,
        bool includeGraphs);
}
