namespace Osadka.Core.Units
{
    public enum Unit { Millimeter, Centimeter, Decimeter, Meter }

    public static class UnitConverter
    {
        // БАЗА = миллиметры

        // Из мм -> в выбранную единицу
        public static double MmTo(double mm, Unit target) => target switch
        {
            Unit.Millimeter => mm,
            Unit.Centimeter => mm / 10.0,
            Unit.Decimeter => mm / 100.0,
            Unit.Meter => mm / 1000.0,
            _ => mm
        };

        // Из выбранной единицы -> в мм
        public static double ToMm(double value, Unit source) => source switch
        {
            Unit.Millimeter => value,
            Unit.Centimeter => value * 10.0,
            Unit.Decimeter => value * 100.0,
            Unit.Meter => value * 1000.0,
            _ => value
        };

        public static string FormatMm(double mm, Unit target, int decimals = 3)
            => $"{Math.Round(MmTo(mm, target), decimals)} " + (target switch
            {
                Unit.Millimeter => "мм",
                Unit.Centimeter => "см",
                Unit.Decimeter => "дм",
                Unit.Meter => "м",
                _ => "мм"
            });
    }
}
