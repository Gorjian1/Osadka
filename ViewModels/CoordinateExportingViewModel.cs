using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using CommunityToolkit.Mvvm.Messaging;
using Microsoft.Win32;
using Osadka.Messages;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using CadAttrib = ACadSharp.Entities.AttributeEntity;
using CadCircle = ACadSharp.Entities.Circle;
using CadInsert = ACadSharp.Entities.Insert;
using CadLine = ACadSharp.Entities.Line;
using CadMText = ACadSharp.Entities.MText;
using CadPolyline = ACadSharp.Entities.LwPolyline;
using CadText = ACadSharp.Entities.TextEntity;
using CadColor = ACadSharp.Color;
using CadLayer = ACadSharp.Tables.Layer;
using WpfPoint = System.Windows.Point;
using WpfColor = System.Windows.Media.Color;
using WpfBrush = System.Windows.Media.Brush;
using WpfSolidColorBrush = System.Windows.Media.SolidColorBrush;

namespace Osadka.ViewModels
{
    public partial class CoordinateExportingViewModel : ObservableObject
    {
        private readonly RawDataViewModel _raw;

        public CoordinateExportingViewModel(RawDataViewModel raw)
        {
            _raw = raw;

            OpenDwgCommand = new RelayCommand(OpenDwg);
            ToggleMeasureCmd = new RelayCommand(() => { IsMeasureMode = !IsMeasureMode; if (IsMeasureMode) IsSelectMode = false; });
            ToggleSelectCmd = new RelayCommand(() => { IsSelectMode = !IsSelectMode; if (IsSelectMode) IsMeasureMode = false; if (!IsSelectMode) SelectedPoints.Clear(); });
            ExportCoordsCmd = new RelayCommand(ExportCoords);
            SendToDataCmd = new RelayCommand(SendToData);
        }

        partial void OnGridStepChanged(double value)
        {
            RebuildGridBrush();
        }

        private DrawingImage _cadDrawing;
        public DrawingImage CadDrawing
        {
            get => _cadDrawing;
            set => SetProperty(ref _cadDrawing, value);
        }

        public CadDocument Doc { get; private set; }

        public ObservableCollection<WpfPoint> SelectedPoints { get; } = new();
        public ObservableCollection<PointCollection> Contours { get; } = new();

        private double _zoomFactor = 1.0;
        public double ZoomFactor
        {
            get => _zoomFactor;
            set
            {
                if (SetProperty(ref _zoomFactor, value))
                {
                    OnPropertyChanged(nameof(EffectiveScale));
                    if (ConstantScreenThickness) RebuildScene();
                }
            }
        }

        public double FitScale { get; private set; } = 1.0;
        public double EffectiveScale => FitScale * ZoomFactor;

        public double PixelsPerUnit { get; private set; }
        public double MinX { get; private set; }
        public double MinY { get; private set; }
        public double MaxX { get; private set; }
        public double MaxY { get; private set; }

        public double ImageWidthPx => (MaxX - MinX) * PixelsPerUnit;
        public double ImageHeightPx => (MaxY - MinY) * PixelsPerUnit;

        public double ScaleBarPixels => CadDrawing == null ? 0 : 100 * PixelsPerUnit;
        public string ScaleBarLabel => CadDrawing == null ? string.Empty : "100 м";

        [ObservableProperty] private bool isMeasureMode;
        [ObservableProperty] private bool isSelectMode;

        [ObservableProperty] private double gridStep = 10.0;

        [ObservableProperty] private bool gridVisible = false;

        [ObservableProperty] private double isoStep = 5.0;

        [ObservableProperty] private bool constantScreenThickness = true;
        [ObservableProperty] private double baseStrokePx = 1.2;

        [ObservableProperty] private bool whiteHaloEnabled = true;
        [ObservableProperty] private double haloExtraPx = 1.0;

        private WpfBrush _gridBrush;
        public WpfBrush GridBrush
        {
            get => _gridBrush;
            private set => SetProperty(ref _gridBrush, value);
        }

        public class LayerVm : ObservableObject
        {
            public string Name { get; }
            private bool _isVisible = true;
            public bool IsVisible { get => _isVisible; set => SetProperty(ref _isVisible, value); }
            public LayerVm(string name, bool visible = true) { Name = name; IsVisible = visible; }
        }
        public ObservableCollection<LayerVm> Layers { get; } = new();

        public IRelayCommand OpenDwgCommand { get; }
        public IRelayCommand ToggleMeasureCmd { get; }
        public IRelayCommand ToggleSelectCmd { get; }
        public IRelayCommand ExportCoordsCmd { get; }
        public IRelayCommand SendToDataCmd { get; }

        private readonly Dictionary<string, List<(Geometry geo, WpfBrush brush)>> _layerCache = new();
        private readonly List<(string layer, string text, double x, double y, double height, WpfBrush brush)> _textCache = new();

        public void OpenDwgFromPath(string filePath)
        {
            try
            {
                Doc = DwgReader.Read(filePath);
                ComputeBounds(Doc, out double minX, out double minY, out double maxX, out double maxY);
                MinX = minX; MinY = minY; MaxX = maxX; MaxY = maxY;

                int targetPx = 3000;
                double dwgW = MaxX - MinX, dwgH = MaxY - MinY;
                PixelsPerUnit = targetPx / Math.Max(dwgW, dwgH);

                RebuildGridBrush();

                RebuildLayersFromDoc();
                BuildGeometryCache();
                RebuildScene();

                ZoomFactor = 1.0;
                OnPropertyChanged(nameof(ScaleBarPixels));
                SelectedPoints.Clear();
                Contours.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка чтения DWG: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenDwg()
        {
            var dlg = new OpenFileDialog
            {
                Filter = "DWG (*.dwg)|*.dwg|DXF (*.dxf)|*.dxf|All files (*.*)|*.*",
                Title = "Открыть чертёж"
            };
            if (dlg.ShowDialog() == true)
            {
                OpenDwgFromPath(dlg.FileName);
            }
        }

        private void RebuildLayersFromDoc()
        {
            Layers.Clear();
            if (Doc?.Layers == null) return;
            foreach (var l in Doc.Layers)
                Layers.Add(new LayerVm(l.Name, true));
        }

        private static readonly Dictionary<int, WpfColor> _aci = new()
        {
            [1] = WpfColor.FromRgb(255, 0, 0),
            [2] = WpfColor.FromRgb(255, 255, 0),
            [3] = WpfColor.FromRgb(0, 255, 0),
            [4] = WpfColor.FromRgb(0, 255, 255),
            [5] = WpfColor.FromRgb(0, 0, 255),
            [6] = WpfColor.FromRgb(255, 0, 255),
            [7] = WpfColor.FromRgb(255, 255, 255),
            [8] = WpfColor.FromRgb(128, 128, 128),
            [9] = WpfColor.FromRgb(255, 128, 0),
        };

        private static WpfBrush Frozen(WpfColor c)
        {
            var b = new WpfSolidColorBrush(c);
            b.Freeze();
            return b;
        }

        private static WpfColor HsvToRgb(double h, double s, double v)
        {
            double C = v * s;
            double X = C * (1 - Math.Abs((h / 60.0) % 2 - 1));
            double m = v - C;
            double r = 0, g = 0, b = 0;
            if (h < 60) { r = C; g = X; b = 0; }
            else if (h < 120) { r = X; g = C; b = 0; }
            else if (h < 180) { r = 0; g = C; b = X; }
            else if (h < 240) { r = 0; g = X; b = C; }
            else if (h < 300) { r = X; g = 0; b = C; }
            else { r = C; g = 0; b = X; }
            return WpfColor.FromRgb(
                (byte)Math.Round((r + m) * 255),
                (byte)Math.Round((g + m) * 255),
                (byte)Math.Round((b + m) * 255));
        }

        private static WpfBrush AciToBrush(int index)
        {
            if (_aci.TryGetValue(index, out var c)) return Frozen(c);
            var hue = (index * 137) % 360;
            return Frozen(HsvToRgb(hue, 0.55, 0.92));
        }

        private WpfBrush GetBrushFor(Entity ent)
        {
            CadColor c = ent.Color;

            if (c.IsTrueColor)
                return Frozen(WpfColor.FromRgb((byte)c.R, (byte)c.G, (byte)c.B));

            if (c.IsByLayer && ent.Layer is CadLayer layer)
            {
                CadColor lc = layer.Color;
                if (lc.IsTrueColor)
                    return Frozen(WpfColor.FromRgb((byte)lc.R, (byte)lc.G, (byte)lc.B));
                if (!lc.IsByLayer)
                    return AciToBrush(lc.Index);
            }
            if (!c.IsTrueColor)
                return AciToBrush(c.Index);
            string key = ent.Layer?.Name ?? "0";
            var hue2 = (uint)key.GetHashCode() % 360u;
            return Frozen(HsvToRgb(hue2, 0.55, 0.92));
        }

        private static bool IsNearWhite(WpfColor col) =>
            col.R > 245 && col.G > 245 && col.B > 245;

        private void BuildGeometryCache()
        {
            _layerCache.Clear();
            _textCache.Clear();
            if (Doc == null) return;

            void AddGeo(string layer, Geometry g, WpfBrush brush)
            {
                if (!_layerCache.TryGetValue(layer, out var list))
                    _layerCache[layer] = list = new List<(Geometry, WpfBrush)>();
                list.Add((g, brush));
            }

            void EmitEntity(Entity ent, string layerOverride = null)
            {
                string layer = layerOverride ?? (ent.Layer?.Name ?? "0");
                var brush = GetBrushFor(ent);

                switch (ent)
                {
                    case CadLine ln:
                        {
                            var g = new StreamGeometry();
                            using (var sg = g.Open())
                            {
                                sg.BeginFigure(new System.Windows.Point(ln.StartPoint.X, ln.StartPoint.Y), false, false);
                                sg.LineTo(new System.Windows.Point(ln.EndPoint.X, ln.EndPoint.Y), true, true);
                            }
                            g.Freeze();
                            AddGeo(layer, g, brush);
                            break;
                        }
                    case CadCircle c:
                        {
                            var g = new EllipseGeometry(new System.Windows.Point(c.Center.X, c.Center.Y), c.Radius, c.Radius);
                            g.Freeze();
                            AddGeo(layer, g, brush);
                            break;
                        }
                    case CadPolyline pl when pl.Vertices.Count > 0:
                        {
                            var g = new StreamGeometry();
                            using (var sg = g.Open())
                            {
                                var v0 = pl.Vertices[0].Location;
                                sg.BeginFigure(new System.Windows.Point(v0.X, v0.Y), false, pl.IsClosed);
                                foreach (var v in pl.Vertices.Skip(1))
                                    sg.LineTo(new System.Windows.Point(v.Location.X, v.Location.Y), true, true);
                            }
                            g.Freeze();
                            AddGeo(layer, g, brush);
                            break;
                        }
                    case CadText txt:
                        _textCache.Add((layer, txt.Value, txt.InsertPoint.X, txt.InsertPoint.Y, txt.Height, brush));
                        break;

                    case CadMText mt:
                        _textCache.Add((layer, mt.Value, mt.InsertPoint.X, mt.InsertPoint.Y, mt.Height, brush));
                        break;

                    case CadInsert ins:
                        if (ins.Block != null)
                        {
                            foreach (var sub in ins.Block.Entities)
                                EmitEntity(sub, ins.Layer?.Name ?? layer);
                        }
                        foreach (CadAttrib at in ins.Attributes)
                            _textCache.Add((layer, at.Value, at.InsertPoint.X, at.InsertPoint.Y, at.Height, brush));
                        break;

                    default:
                        break;
                }
            }

            foreach (var ent in Doc.Entities)
                EmitEntity(ent);
        }

        private void RebuildScene()
        {
            if (Doc == null) return;

            var visible = new HashSet<string>(Layers.Where(l => l.IsVisible).Select(l => l.Name));

            double strokeWorld = BaseStrokePx / PixelsPerUnit;
            if (ConstantScreenThickness)
                strokeWorld /= Math.Max(EffectiveScale, 1e-6);

            var group = new DrawingGroup();
            var m = new Matrix(PixelsPerUnit, 0, 0, -PixelsPerUnit, -MinX * PixelsPerUnit, MaxY * PixelsPerUnit);
            group.Transform = new MatrixTransform(m);

            using (var dc = group.Open())
            {
                foreach (var kv in _layerCache)
                {
                    if (!visible.Contains(kv.Key)) continue;

                    foreach (var item in kv.Value)
                    {
                        var geo = item.geo;
                        var brush = item.brush;

                        var pen = new Pen(brush, strokeWorld); pen.Freeze();

                        if (WhiteHaloEnabled && brush is WpfSolidColorBrush scb && IsNearWhite(scb.Color))
                        {
                            double haloWorld = strokeWorld + ScreenPxToWorld(HaloExtraPx);
                            var haloPen = new Pen(Brushes.Black, haloWorld); haloPen.Freeze();
                            dc.DrawGeometry(null, haloPen, geo);
                        }

                        dc.DrawGeometry(null, pen, geo);
                    }
                }

                foreach (var t in _textCache)
                {
                    if (!visible.Contains(t.layer)) continue;
                    DrawCadText(dc, t.text, t.x, t.y, t.height, t.brush);
                }
            }

            group.Freeze();
            CadDrawing = new DrawingImage(group);
            CadDrawing.Freeze();

            OnPropertyChanged(nameof(ImageWidthPx));
            OnPropertyChanged(nameof(ImageHeightPx));
        }

        private void RebuildGridBrush()
        {
            if (PixelsPerUnit <= 0 || GridStep <= 0)
            {
                GridBrush = null;
                return;
            }

            double cellPx = GridStep * PixelsPerUnit;

            var dg = new DrawingGroup();
            using (var dc = dg.Open())
            {
                var pen = new Pen(new SolidColorBrush(System.Windows.Media.Color.FromArgb(90, 160, 160, 160)), 1.0);
                pen.Freeze();
                var rect = new System.Windows.Rect(0, 0, cellPx, cellPx);
                dc.DrawRectangle(null, pen, rect);
            }
            dg.Freeze();

            var brush = new DrawingBrush(dg)
            {
                TileMode = TileMode.Tile,
                ViewportUnits = BrushMappingMode.Absolute,
                Viewport = new System.Windows.Rect(0, 0, cellPx, cellPx),
                Stretch = Stretch.None
            };
            brush.Freeze();
            GridBrush = brush;
        }

        private static void DrawCadText(DrawingContext dc, string txt, double x, double y, double height, WpfBrush brush)
        {
            double fontUnits = Math.Max(height, 0.5);
            var ft = new FormattedText(
                txt ?? string.Empty,
                System.Globalization.CultureInfo.CurrentUICulture,
                FlowDirection.LeftToRight,
                new Typeface("Segoe UI"),
                fontUnits,
                brush,
                1.0);

            dc.PushTransform(new TranslateTransform(x, y));
            dc.PushTransform(new ScaleTransform(1, -1));

            var origin = new System.Windows.Point(0, -ft.Baseline);
            dc.DrawText(ft, origin);

            dc.Pop();
            dc.Pop();
        }

        private static void ComputeBounds(CadDocument doc, out double minX, out double minY, out double maxX, out double maxY)
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

                    case CadPolyline pl when pl.Vertices.Count > 0:
                        foreach (var v in pl.Vertices) upd(v.Location.X, v.Location.Y);
                        break;

                    case CadText t:
                        upd(t.InsertPoint.X, t.InsertPoint.Y);
                        break;

                    case CadMText mt:
                        upd(mt.InsertPoint.X, mt.InsertPoint.Y);
                        break;

                    case CadInsert ins:
                        if (ins.Block != null)
                        {
                            foreach (var sub in ins.Block.Entities)
                            {
                                switch (sub)
                                {
                                    case CadLine l2:
                                        upd(l2.StartPoint.X, l2.StartPoint.Y);
                                        upd(l2.EndPoint.X, l2.EndPoint.Y);
                                        break;
                                    case CadCircle c2:
                                        upd(c2.Center.X - c2.Radius, c2.Center.Y - c2.Radius);
                                        upd(c2.Center.X + c2.Radius, c2.Center.Y + c2.Radius);
                                        break;
                                    case CadPolyline pl2 when pl2.Vertices.Count > 0:
                                        foreach (var v in pl2.Vertices) upd(v.Location.X, v.Location.Y);
                                        break;
                                }
                            }
                        }
                        foreach (CadAttrib at in ins.Attributes)
                            upd(at.InsertPoint.X, at.InsertPoint.Y);
                        break;
                }
            }

            if (!double.IsFinite(mnX) || !double.IsFinite(mnY) ||
                !double.IsFinite(mxX) || !double.IsFinite(mxY))
            {
                mnX = mnY = 0; mxX = mxY = 1;
            }

            minX = mnX; minY = mnY; maxX = mxX; maxY = mxY;
        }

        double ScreenPxToWorld(double px)
        {
            double w = px / PixelsPerUnit;
            if (ConstantScreenThickness)
                w /= Math.Max(EffectiveScale, 1e-6);
            return w;
        }

        public void FitToViewport(double viewportW, double viewportH)
        {
            if (CadDrawing == null || viewportW <= 0 || viewportH <= 0) return;

            double worldW = (MaxX - MinX);
            double worldH = (MaxY - MinY);
            if (worldW <= 0 || worldH <= 0) return;

            double scaleX = viewportW / (worldW * PixelsPerUnit);
            double scaleY = viewportH / (worldH * PixelsPerUnit);
            FitScale = Math.Min(scaleX, scaleY);
            OnPropertyChanged(nameof(EffectiveScale));
            OnPropertyChanged(nameof(ScaleBarPixels));
        }

        private void ExportCoords()
        {
            if (SelectedPoints.Count == 0)
            {
                MessageBox.Show("Нечего экспортировать", "Экспорт",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dlg = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv" };
            if (dlg.ShowDialog() != true) return;

            using var writer = new StreamWriter(dlg.FileName);
            writer.WriteLine("X;Y");
            foreach (var pt in SelectedPoints)
                writer.WriteLine($"{pt.X:F2};{pt.Y:F2}");
        }

        private void SendToData()
        {
            var scaled = SelectedPoints
                .Select(p => new System.Windows.Point(
                    Math.Round(p.X * _raw.CoordScale, 2),
                    Math.Round(p.Y * _raw.CoordScale, 2)))
                .ToList();

            WeakReferenceMessenger.Default.Send(new CoordinatesMessage(scaled));
        }
    }
}
