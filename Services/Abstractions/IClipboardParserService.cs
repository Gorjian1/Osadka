using Osadka.Core.Units;
using Osadka.Models;
using System.Collections.Generic;

namespace Osadka.Services.Abstractions;

/// <summary>
/// Сервис для парсинга данных из буфера обмена
/// </summary>
public interface IClipboardParserService
{
    /// <summary>
    /// Результат парсинга буфера обмена
    /// </summary>
    public class ParseResult
    {
        /// <summary>
        /// Тип распознанных данных
        /// </summary>
        public enum DataType
        {
            None,           // Нераспознанный формат
            Ids,            // 1 колонка: ID точек
            Coordinates,    // 2 колонки: X, Y координаты
            Measurements3,  // 3 колонки: Отметка, Осадка, Общая
            Measurements4   // 4 колонки: Отметка, Осадка, Общая, ID
        }

        public DataType Type { get; set; } = DataType.None;
        public List<string> Ids { get; set; } = new();
        public List<CoordRow> Coordinates { get; set; } = new();
        public List<MeasurementRow> Measurements { get; set; } = new();
    }

    /// <summary>
    /// Парсит содержимое буфера обмена
    /// </summary>
    /// <param name="clipboardText">Текст из буфера обмена</param>
    /// <param name="cycleNumber">Номер цикла для измерений</param>
    /// <param name="existingIds">Существующие ID точек (для 3-колоночного формата)</param>
    /// <param name="coordUnit">Единица измерения координат</param>
    /// <returns>Результат парсинга</returns>
    ParseResult Parse(
        string clipboardText,
        int cycleNumber,
        IReadOnlyList<string> existingIds,
        Unit coordUnit);
}
