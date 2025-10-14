using System.Windows;
using System.Windows.Media;

namespace Osadka.UI.Cad
{
    public static class MatrixExtensions
    {
        public static Point Transform(this Matrix m, Point p) => m.Transform(p);

        public static Matrix Translate(this Matrix m, double dx, double dy)
        {
            m.Translate(dx, dy);
            return m;
        }

        public static Matrix Scale(this Matrix m, double sx, double sy)
        {
            m.Scale(sx, sy);
            return m;
        }
    }
}
