using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using CommunityToolkit.Mvvm.Messaging;
using Microsoft.Win32;
using Osadka.Messages;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using CadAttrib = ACadSharp.Entities.AttributeEntity;
using CadCircle = ACadSharp.Entities.Circle;
using CadInsert = ACadSharp.Entities.Insert;
using CadLine = ACadSharp.Entities.Line;
using CadMText = ACadSharp.Entities.MText;
using CadPolyline = ACadSharp.Entities.LwPolyline;
using CadText = ACadSharp.Entities.TextEntity;
using WpfPoint = System.Windows.Point;

namespace Osadka.ViewModels
{
    public partial class CoordinateExportingViewModel : ObservableObject
    {

        private BitmapSource _cadBitmap;
        private readonly RawDataViewModel _raw;
        public BitmapSource CadBitmap
        {
            get => _cadBitmap;
            set => SetProperty(ref _cadBitmap, value);
        }

        private double _zoomFactor = 1.0;
        public double ZoomFactor
        {
            get => _zoomFactor;
            set
            {
                if (SetProperty(ref _zoomFactor, value))
                    OnPropertyChanged(nameof(EffectiveScale));
            }
        }

        public double FitScale { get; private set; } = 1.0;
        public double EffectiveScale => FitScale * ZoomFactor;

        public double PixelsPerUnit { get; private set; }
        public double ScaleBarPixels =>
            CadBitmap == null ? 0 : 100 * PixelsPerUnit * EffectiveScale;
        public string ScaleBarLabel =>
            CadBitmap == null ? string.Empty : "100 м";

        public double MinX { get; private set; }
        public double MinY { get; private set; }
        public double MaxX { get; private set; }
        public double MaxY { get; private set; }

        [ObservableProperty] private bool isMeasureMode;
        [ObservableProperty] private bool isSelectMode;



        [ObservableProperty] private double gridStep = 10.0;
        [ObservableProperty] private double isoStep = 5.0;



        public ObservableCollection<WpfPoint> SelectedPoints { get; } = new();
        public ObservableCollection<PointCollection> Contours { get; } = new();

        public IRelayCommand OpenDwgCommand { get; }
        public IRelayCommand ToggleMeasureCmd { get; }
        public IRelayCommand ToggleSelectCmd { get; }
        public IRelayCommand ExportCoordsCmd { get; }
        public IRelayCommand SendToDataCmd { get; }
        public IRelayCommand BuildMapCmd { get; }


        public CoordinateExportingViewModel(RawDataViewModel raw)
        {
            _raw = raw;

            OpenDwgCommand = new RelayCommand(OpenDwg);
            ExportCoordsCmd = new RelayCommand(ExportCoords);
            SendToDataCmd = new RelayCommand(SendToData);
            ToggleMeasureCmd = new RelayCommand(() =>
            {
                IsMeasureMode = !IsMeasureMode;
                if (IsMeasureMode) IsSelectMode = false;
            });
            ToggleSelectCmd = new RelayCommand(() =>
            {
                IsSelectMode = !IsSelectMode;
                if (IsSelectMode) IsMeasureMode = false;
                if (!IsSelectMode) SelectedPoints.Clear();
            });

            BuildMapCmd = new RelayCommand(BuildMap,
                () => SelectedPoints.Count >= 3 &&
                      _raw.CoordRows.Count >= SelectedPoints.Count &&
                      _raw.DataRows.Count >= SelectedPoints.Count);

            SelectedPoints.CollectionChanged += (_, __) => BuildMapCmd.NotifyCanExecuteChanged();
        }


        private void BuildMap()
        {
            var count = SelectedPoints.Count;
            var triples = Enumerable.Range(0, count)
                .Select(i => (
                    X: SelectedPoints[i].X,
                    Y: SelectedPoints[i].Y,
                    Mark: _raw.DataRows[i].Mark))
                .Where(t => t.Mark.HasValue)
                .Select(t => (t.X, t.Y, Z: t.Mark.Value))
                .ToList();
            if (triples.Count < 3)
            {
                MessageBox.Show("Недостаточно точек для изолиний (нужно ≥3).", "Карта осадков", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var isolines = ConrecHelper.Generate(triples, IsoStep, GridStep);

            Contours.Clear();
            foreach (var line in isolines)
            {
                var pc = new PointCollection();
                foreach (var (x, y) in line)
                {
                    var px = (x - MinX) * PixelsPerUnit;
                    var py = (MaxY - y) * PixelsPerUnit;
                    pc.Add(new WpfPoint(px, py));
                }
                if (pc.Count > 1) Contours.Add(pc);
            }
        }


        private void SendToData()
        {
            var msg = new CoordinatesMessage(this.SelectedPoints);
            WeakReferenceMessenger.Default.Send(msg);
        }

        public void FitToViewport(double viewportW, double viewportH)
        {
            if (CadBitmap == null || viewportW <= 0 || viewportH <= 0) return;
            FitScale = Math.Min(
                viewportW / CadBitmap.PixelWidth,
                viewportH / CadBitmap.PixelHeight);
            OnPropertyChanged(nameof(FitScale));
            OnPropertyChanged(nameof(EffectiveScale));
            OnPropertyChanged(nameof(ScaleBarPixels));
        }

        private void OpenDwg()
        {
            var dlg = new OpenFileDialog { Filter = "AutoCAD DWG|*.dwg" };
            if (dlg.ShowDialog() != true) return;

            try
            {
                var doc = DwgReader.Read(dlg.FileName);
                CadBitmap = RenderCadToBitmap(doc, 3000, out double ppu, out var bounds);
                PixelsPerUnit = ppu;
                (MinX, MinY, MaxX, MaxY) = bounds;
                ZoomFactor = 1.0;
                OnPropertyChanged(nameof(ScaleBarPixels));
                SelectedPoints.Clear();
                Contours.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка чтения DWG: {ex.Message}",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void ExportCoords()
        {
            if (SelectedPoints.Count == 0)
            {
                MessageBox.Show(
                    "Нечего экспортировать",
                    "Экспорт",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            var dlg = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv" };
            if (dlg.ShowDialog() != true) return;

            using var writer = new StreamWriter(dlg.FileName);
            writer.WriteLine("X;Y");
            foreach (var pt in SelectedPoints)
                writer.WriteLine($"{pt.X:F3};{pt.Y:F3}");
        }


        private static BitmapSource RenderCadToBitmap(
            CadDocument doc,
            int targetPx,
            out double pxPerUnit,
            out (double minX, double minY, double maxX, double maxY) bounds)
        {
            GetBounds(doc, out double minX, out double minY, out double maxX, out double maxY);
            bounds = (minX, minY, maxX, maxY);

            double dwgW = maxX - minX, dwgH = maxY - minY;
            double scale = targetPx / Math.Max(dwgW, dwgH);
            pxPerUnit = scale;

            int bmpW = (int)Math.Ceiling(dwgW * scale);
            int bmpH = (int)Math.Ceiling(dwgH * scale);

            var dv = new DrawingVisual();
            using (var dc = dv.RenderOpen())
            {
                foreach (var ent in doc.Entities)
                {
                    DrawEntityOrText(dc, ent, minX, minY, scale, bmpH);
                }
            }

            var bmp = new RenderTargetBitmap(bmpW, bmpH, 96, 96, PixelFormats.Pbgra32);
            bmp.Render(dv);
            bmp.Freeze();
            return bmp;
        }

        private static void DrawEntityOrText(
            DrawingContext dc,
            Entity ent,
            double minX, double minY,
            double scale, int bmpH)
        {
            var pen = new Pen(Brushes.Black, 1);

            switch (ent)
            {
                case CadLine ln:
                    dc.DrawLine(pen,
                        ToPx(ln.StartPoint.X, ln.StartPoint.Y, minX, minY, scale, bmpH),
                        ToPx(ln.EndPoint.X, ln.EndPoint.Y, minX, minY, scale, bmpH));
                    break;

                case CadCircle c:
                    dc.DrawEllipse(null, pen,
                        ToPx(c.Center.X, c.Center.Y, minX, minY, scale, bmpH),
                        c.Radius * scale, c.Radius * scale);
                    break;

                case CadPolyline pl when pl.Vertices.Count > 0:
                    var geo = new StreamGeometry();
                    using (var sg = geo.Open())
                    {
                        var v0 = pl.Vertices[0].Location;
                        sg.BeginFigure(
                            ToPx(v0.X, v0.Y, minX, minY, scale, bmpH),
                            false, pl.IsClosed);
                        foreach (var v in pl.Vertices.Skip(1))
                            sg.LineTo(
                                ToPx(v.Location.X, v.Location.Y, minX, minY, scale, bmpH),
                                true, true);
                    }
                    geo.Freeze();
                    dc.DrawGeometry(null, pen, geo);
                    break;

                case CadText txt:
                    DrawCadText(dc, txt.Value,
                        txt.InsertPoint.X, txt.InsertPoint.Y, txt.Height,
                        minX, minY, scale, bmpH, Brushes.Black);
                    break;

                case CadMText mt:
                    DrawCadText(dc, mt.Value,
                        mt.InsertPoint.X, mt.InsertPoint.Y, mt.Height,
                        minX, minY, scale, bmpH, Brushes.Black);
                    break;

                case CadInsert ins:
                    if (ins.Block != null)
                        foreach (var sub in ins.Block.Entities)
                            DrawEntityOrText(dc, sub, minX, minY, scale, bmpH);
                    foreach (CadAttrib at in ins.Attributes)
                    {
                        DrawCadText(dc, at.Value,
                            at.InsertPoint.X, at.InsertPoint.Y, at.Height,
                            minX, minY, scale, bmpH, Brushes.Black);
                    }
                    break;
            }
        }

        private static void GetBounds(
            CadDocument doc,
            out double minX, out double minY,
            out double maxX, out double maxY)
        {
            double mnX = double.PositiveInfinity, mnY = double.PositiveInfinity;
            double mxX = double.NegativeInfinity, mxY = double.NegativeInfinity;
            void upd(double x, double y)
            {
                if (x < mnX) mnX = x;
                if (y < mnY) mnY = y;
                if (x > mxX) mxX = x;
                if (y > mxY) mxY = y;
            }

            foreach (var e in doc.Entities)
            {
                switch (e)
                {
                    case CadLine ln:
                        upd(ln.StartPoint.X, ln.StartPoint.Y);
                        upd(ln.EndPoint.X, ln.EndPoint.Y);
                        break;
                    case CadCircle c:
                        upd(c.Center.X - c.Radius, c.Center.Y - c.Radius);
                        upd(c.Center.X + c.Radius, c.Center.Y + c.Radius);
                        break;
                    case CadPolyline pl:
                        foreach (var v in pl.Vertices)
                            upd(v.Location.X, v.Location.Y);
                        break;
                    case CadInsert ins:
                        upd(ins.InsertPoint.X, ins.InsertPoint.Y);
                        if (ins.Block != null)
                            foreach (var sub in ins.Block.Entities)
                                if (sub is CadLine l2)
                                {
                                    upd(l2.StartPoint.X, l2.StartPoint.Y);
                                    upd(l2.EndPoint.X, l2.EndPoint.Y);
                                }
                                else if (sub is CadCircle c2)
                                {
                                    upd(c2.Center.X - c2.Radius, c2.Center.Y - c2.Radius);
                                    upd(c2.Center.X + c2.Radius, c2.Center.Y + c2.Radius);
                                }
                                else if (sub is CadPolyline pl2)
                                {
                                    foreach (var v2 in pl2.Vertices)
                                        upd(v2.Location.X, v2.Location.Y);
                                }
                        break;
                }
            }

            minX = mnX; minY = mnY;
            maxX = mxX; maxY = mxY;
        }
        private static WpfPoint ToPx(
            double x, double y,
            double minX, double minY,
            double scale, int bmpH) =>
            new WpfPoint((x - minX) * scale,
                         bmpH - (y - minY) * scale);

        private static void DrawCadText(
            DrawingContext dc,
            string txt,
            double x, double y,
            double height,
            double minX, double minY,
            double scale, int bmpH,
            Brush brush)
        {
            double fontPx = Math.Max(height, 0.5) * scale;
            var ft = new FormattedText(
                txt,
                System.Globalization.CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                new Typeface("Arial"),
                fontPx,
                brush,
                1.0);

            var p = new WpfPoint(
                (x - minX) * scale,
                bmpH - (y - minY) * scale - ft.Height);

            dc.DrawText(ft, p);
        }
    }

    internal static class ConrecHelper
    {
        public static IEnumerable<IEnumerable<(double X, double Y)>> Generate(
            IReadOnlyList<(double X, double Y, double Z)> pts,
            double isoStep,
            double cellSize)
        {
            if (pts.Count < 3) yield break;

            double minX = pts.Min(p => p.X), maxX = pts.Max(p => p.X);
            double minY = pts.Min(p => p.Y), maxY = pts.Max(p => p.Y);

            int nx = Math.Max(2, (int)Math.Ceiling((maxX - minX) / cellSize));
            int ny = Math.Max(2, (int)Math.Ceiling((maxY - minY) / cellSize));

            var grid = new double[nx + 1, ny + 1];
            for (int i = 0; i <= nx; i++)
            {
                for (int j = 0; j <= ny; j++)
                {
                    double x = minX + i * cellSize;
                    double y = minY + j * cellSize;
                    grid[i, j] = Idw(x, y, pts);
                }
            }

            double zMin = grid.Cast<double>().Min();
            double zMax = grid.Cast<double>().Max();
            for (double level = Math.Ceiling(zMin / isoStep) * isoStep; level <= zMax; level += isoStep)
            {
                foreach (var poly in March(level, grid, minX, minY, cellSize))
                    yield return poly;
            }
        }

        private static double Idw(double x, double y, IReadOnlyList<(double X, double Y, double Z)> pts, double power = 2)
        {
            const double eps = 1e-12;
            double num = 0, den = 0;
            foreach (var (px, py, z) in pts)
            {
                double d2 = (px - x) * (px - x) + (py - y) * (py - y) + eps;
                double w = 1 / Math.Pow(d2, power / 2);
                num += w * z; den += w;
            }
            return num / den;
        }

        private static readonly (int dx, int dy)[] Corner =
        {
            (0,0),(1,0),(1,1),(0,1)
        };

        private static IEnumerable<IEnumerable<(double X, double Y)>> March(
            double level,
            double[,] g,
            double x0,
            double y0,
            double cell)
        {
            int nx = g.GetLength(0) - 1;
            int ny = g.GetLength(1) - 1;

            for (int i = 0; i < nx; i++)
            {
                for (int j = 0; j < ny; j++)
                {
                    int mask = 0;
                    if (g[i, j] > level) mask |= 1;
                    if (g[i + 1, j] > level) mask |= 2;
                    if (g[i + 1, j + 1] > level) mask |= 4;
                    if (g[i, j + 1] > level) mask |= 8;
                    if (mask == 0 || mask == 15) continue;

                    var poly = new List<(double, double)>();
                    for (int k = 0; k < 4; k++)
                    {
                        int k1 = k, k2 = (k + 1) % 4;
                        bool a = (mask & (1 << k1)) != 0;
                        bool b = (mask & (1 << k2)) != 0;
                        if (a == b) continue;

                        var (dx1, dy1) = Corner[k1];
                        var (dx2, dy2) = Corner[k2];

                        double z1 = g[i + dx1, j + dy1];
                        double z2 = g[i + dx2, j + dy2];
                        double t = (level - z1) / (z2 - z1 + double.Epsilon);

                        double x = x0 + (i + dx1 + t * (dx2 - dx1)) * cell;
                        double y = y0 + (j + dy1 + t * (dy2 - dy1)) * cell;
                        poly.Add((x, y));
                    }
                    if (poly.Count > 1) yield return poly;
                }
            }
        }
    }
}