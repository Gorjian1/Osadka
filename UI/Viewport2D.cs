using System;
using System.Windows;
using System.Windows.Media;

namespace Osadka.UI.Cad
{
    /// <summary>
    /// Камера/вьюпорт для 2D DWG: хранит зум/пан и строит матрицы.
    /// Внутренние единицы мира = миллиметры.
    /// </summary>
    public sealed class Viewport2D
    {
        public Rect WorldBoundsMm { get; private set; } = Rect.Empty;
        public double PixelsPerMm { get; private set; } = 1.0;

        public double Zoom { get; private set; } = 1.0;      // масштаб камеры
        public Vector Pan { get; private set; } = new Vector(0, 0); // сдвиг камеры в пикселях

        public double MinZoom { get; set; } = 0.05;
        public double MaxZoom { get; set; } = 50.0;

        // Инициализация под размер экрана: вписываем чертёж и центрируем
        public void FitToExtents(Rect worldBoundsMm, Size viewportPixels, double padding = 0.9)
        {
            if (worldBoundsMm.IsEmpty || viewportPixels.Width <= 1 || viewportPixels.Height <= 1)
                return;

            WorldBoundsMm = worldBoundsMm;

            var sx = viewportPixels.Width / worldBoundsMm.Width;
            var sy = viewportPixels.Height / worldBoundsMm.Height;
            PixelsPerMm = Math.Min(sx, sy) * padding;

            Zoom = 1.0;
            // центрируем остаток: пан — в пикселях
            var imgWidth = worldBoundsMm.Width * PixelsPerMm;
            var imgHeight = worldBoundsMm.Height * PixelsPerMm;
            Pan = new Vector(
                (viewportPixels.Width - imgWidth) * 0.5,
                (viewportPixels.Height - imgHeight) * 0.5
            );
        }

        public void ZoomAt(Point screenPointPx, double factor, Size viewportPixels)
        {
            var clamped = Math.Clamp(Zoom * factor, MinZoom, MaxZoom);
            factor = clamped / Zoom;
            if (Math.Abs(factor - 1.0) < 1e-9) return;

            // Мировая точка под курсором до зума:
            var worldBefore = ScreenToWorld(screenPointPx, viewportPixels);

            Zoom = clamped;

            // После изменения масштаба держим ту же мировую точку под курсором:
            var screenAfter = WorldToScreen(worldBefore, viewportPixels);
            var delta = (Vector)(screenPointPx - screenAfter);
            Pan += delta;
        }

        public void PanBy(Vector deltaPx) => Pan += deltaPx;

        // Матрицы:
        // Mfit: world(mm)->pixels без зума/панорамирования; View: зум/пан; FlipY: из Y-up в Y-down
        private Matrix BuildModelToScreen(Size viewportPixels)
        {
            var m = Matrix.Identity;
            // 1) world origin at min corner
            m.Translate(-WorldBoundsMm.Left, -WorldBoundsMm.Top);
            // 2) scale mm->px
            m.Scale(PixelsPerMm, PixelsPerMm);
            // 3) camera (zoom & pan)
            m.Scale(Zoom, Zoom);
            m.Translate(Pan.X, Pan.Y);
            // 4) flip Y to WPF screen
            m.Scale(1, -1);
            m.Translate(0, viewportPixels.Height);
            return m;
        }

        public Point WorldToScreen(Point worldMm, Size viewportPixels)
            => BuildModelToScreen(viewportPixels).Transform(worldMm);

        public Point ScreenToWorld(Point screenPx, Size viewportPixels)
        {
            var m = BuildModelToScreen(viewportPixels);
            m.Invert();
            return m.Transform(screenPx);
        }
    }
}
