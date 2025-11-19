using Osadka.Core.Units;
using Osadka.Models;
using System.Collections.Generic;

namespace Osadka.Services.Abstractions;

/// <summary>
/// Сервис для импорта данных из Excel файлов
/// </summary>
public interface IExcelImportService
{
    /// <summary>
    /// Результат импорта из Excel
    /// </summary>
    public class ImportResult
    {
        /// <summary>
        /// Данные по объектам: ObjectNumber -> (CycleNumber -> List<MeasurementRow>)
        /// </summary>
        public Dictionary<int, Dictionary<int, List<MeasurementRow>>> Objects { get; set; } = new();

        /// <summary>
        /// Заголовки циклов: CycleNumber -> Label
        /// </summary>
        public Dictionary<int, string> CycleHeaders { get; set; } = new();

        /// <summary>
        /// Рекомендуемый номер объекта для выбора
        /// </summary>
        public int SuggestedObjectNumber { get; set; } = 1;

        /// <summary>
        /// Рекомендуемый номер цикла для выбора
        /// </summary>
        public int SuggestedCycleNumber { get; set; } = 1;
    }

    /// <summary>
    /// Импортирует данные из Excel файла с интерактивным выбором листа/объекта/цикла
    /// </summary>
    /// <param name="filePath">Путь к Excel файлу</param>
    /// <param name="coordUnit">Текущие единицы измерения координат (для конвертации)</param>
    /// <returns>Результат импорта или null если пользователь отменил</returns>
    ImportResult? ImportFromExcel(string filePath, Unit coordUnit);
}
