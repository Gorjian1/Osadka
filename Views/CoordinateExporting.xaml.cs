using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using Osadka.ViewModels;
using WpfPoint = System.Windows.Point;

namespace Osadka.Views
{
    public partial class CoordinateExporting : UserControl
    {
        private CoordinateExportingViewModel Vm => (CoordinateExportingViewModel)DataContext;

        private bool _panning;
        private WpfPoint _panStart;
        private double _hStart, _vStart;

        private WpfPoint? _measureStart;
        private readonly Line _ruler = new() { Stroke = Brushes.Red, StrokeThickness = 1 };
        private readonly TextBlock _label = new() { Foreground = Brushes.Red, FontSize = 12 };

        public CoordinateExporting(RawDataViewModel raw)
        {
            InitializeComponent();

            DataContext = new CoordinateExportingViewModel(raw);

            Loaded += (_, __) => Fit();
            SizeChanged += (_, __) => Fit();
            OverlayCanvas.Children.Add(_ruler);
            OverlayCanvas.Children.Add(_label);
            _ruler.Visibility = _label.Visibility = Visibility.Collapsed;

            Focusable = true;
            Focus();
        }
        private void BtnBuildMap_Click(object sender, RoutedEventArgs e)
        {
            Vm.BuildMapCmd.Execute(null);
        }
        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            Vm.OpenDwgCommand.Execute(null);
            UpdateToolbarButtons();
        }

        private void BtnRuler_Click(object sender, RoutedEventArgs e)
        {
            Vm.IsMeasureMode = !Vm.IsMeasureMode;
            if (Vm.IsMeasureMode)
            {
                Vm.IsSelectMode = false;
                _measureStart = null;
                _ruler.Visibility = _label.Visibility = Visibility.Collapsed;
            }
            UpdateToolbarButtons();
        }

        private void BtnSelect_Click(object sender, RoutedEventArgs e)
        {
            Vm.IsSelectMode = !Vm.IsSelectMode;
            if (Vm.IsSelectMode)
            {
                Vm.IsMeasureMode = false;
                SelectedPointsClearOverlay();
            }
            UpdateToolbarButtons();
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            Vm.ExportCoordsCmd.Execute(null);
        }

        private void UpdateToolbarButtons()
        {
            BtnRuler.Background = Vm.IsMeasureMode ? Brushes.LightBlue : Brushes.Transparent;
            BtnSelect.Background = Vm.IsSelectMode ? Brushes.LightBlue : Brushes.Transparent;
        }

        private void SelectedPointsClearOverlay()
        {
            OverlayCanvas.Children.Clear();
            OverlayCanvas.Children.Add(_ruler);
            OverlayCanvas.Children.Add(_label);
            Vm.SelectedPoints.Clear();
        }
        private void Fit()
        {
            Vm.FitToViewport(Viewport.ActualWidth, Viewport.ActualHeight - 30);
        }

        private void ImageWheel(object sender, MouseWheelEventArgs e)
        {
            const double step = 1.2;
            double old = Vm.ZoomFactor;
            Vm.ZoomFactor *= e.Delta > 0 ? step : 1 / step;
            e.Handled = true;

            var p = e.GetPosition(Viewport);
            double cx = Viewport.HorizontalOffset + p.X;
            double cy = Viewport.VerticalOffset + p.Y;
            double k = Vm.ZoomFactor / old;

            Viewport.ScrollToHorizontalOffset(cx * k - p.X);
            Viewport.ScrollToVerticalOffset(cy * k - p.Y);
        }
        private void Viewport_RightDown(object s, MouseButtonEventArgs e)
        {
            _panning = true;
            _panStart = e.GetPosition(Viewport);
            _hStart = Viewport.HorizontalOffset;
            _vStart = Viewport.VerticalOffset;
            Viewport.Cursor = Cursors.SizeAll;
            Viewport.CaptureMouse();
            e.Handled = true;
        }

        private void Viewport_Move(object s, MouseEventArgs e)
        {
            if (_panning)
            {
                var cur = e.GetPosition(Viewport);
                Viewport.ScrollToHorizontalOffset(_hStart - (cur.X - _panStart.X));
                Viewport.ScrollToVerticalOffset(_vStart - (cur.Y - _panStart.Y));
            }

            if (Vm.IsMeasureMode && _measureStart is WpfPoint p0)
            {
                UpdateRuler(p0, e.GetPosition(ImageHost));
            }
        }

        private void Viewport_RightUp(object s, MouseButtonEventArgs e)
        {
            if (_panning)
            {
                _panning = false;
                Viewport.ReleaseMouseCapture();
                Viewport.Cursor = Cursors.Arrow;
                e.Handled = true;
            }
        }
        private void Viewport_LeftDown(object s, MouseButtonEventArgs e)
        {
            var p = e.GetPosition(ImageHost);

            if (Vm.IsMeasureMode)
            {
                HandleMeasureClick(p);
            }
            else if (Vm.IsSelectMode)
            {
                var cad = PxToCad(p);
                Vm.SelectedPoints.Add(new WpfPoint(cad.X, cad.Y));

                const double sz = 6;
                var l1 = new Line
                {
                    X1 = p.X - sz,
                    Y1 = p.Y - sz,
                    X2 = p.X + sz,
                    Y2 = p.Y + sz,
                    Stroke = Brushes.Red,
                    StrokeThickness = 2
                };
                var l2 = new Line
                {
                    X1 = p.X - sz,
                    Y1 = p.Y + sz,
                    X2 = p.X + sz,
                    Y2 = p.Y - sz,
                    Stroke = Brushes.Red,
                    StrokeThickness = 2
                };
                OverlayCanvas.Children.Add(l1);
                OverlayCanvas.Children.Add(l2);
            }
        }

        private void UserControl_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && Vm.IsSelectMode)
            {
                Vm.ExportCoordsCmd.Execute(null);
                e.Handled = true;
            }
        }

        private void HandleMeasureClick(WpfPoint p)
        {
            if (_measureStart == null)
            {
                _measureStart = p;
                _ruler.Visibility = _label.Visibility = Visibility.Visible;
            }
            else
            {
                _measureStart = null;
                _ruler.Visibility = _label.Visibility = Visibility.Collapsed;
            }
        }

        private void UpdateRuler(WpfPoint p0, WpfPoint p1)
        {
            _ruler.X1 = p0.X; _ruler.Y1 = p0.Y;
            _ruler.X2 = p1.X; _ruler.Y2 = p1.Y;

            var c0 = PxToCad(p0);
            var c1 = PxToCad(p1);
            double dx = c1.X - c0.X, dy = c1.Y - c0.Y;
            double dist = Math.Sqrt(dx * dx + dy * dy);

            _label.Text = $"{dist:F2}";
            Canvas.SetLeft(_label, (p0.X + p1.X) / 2);
            Canvas.SetTop(_label, (p0.Y + p1.Y) / 2);
        }

        private (double X, double Y) PxToCad(WpfPoint p)
        {
            double imgX = (p.X + Viewport.HorizontalOffset) / Vm.EffectiveScale;
            double imgY = (p.Y + Viewport.VerticalOffset) / Vm.EffectiveScale;

            double dwgX = Vm.MinX + imgX / Vm.PixelsPerUnit;
            double dwgY = Vm.MaxY - imgY / Vm.PixelsPerUnit;
            return (dwgX, dwgY);
        }

    }
}
