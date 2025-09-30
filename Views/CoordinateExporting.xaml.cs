using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Data;
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
        private readonly Stack<UIElement[]> _crossStack = new();

        public CoordinateExporting(RawDataViewModel raw)
        {
            InitializeComponent();
            DataContext = new CoordinateExportingViewModel(raw);

            Loaded += (_, __) =>
            {
                if (!string.IsNullOrWhiteSpace(raw.DrawingPath) && System.IO.File.Exists(raw.DrawingPath))
                    Vm.OpenDwgFromPath(raw.DrawingPath);
                Fit();
            };
            SizeChanged += (_, __) => Fit();

            OverlayCanvas.Children.Add(_ruler);
            OverlayCanvas.Children.Add(_label);
            _ruler.Visibility = _label.Visibility = Visibility.Collapsed;

            BtnLayers.ContextMenu = BuildLayersMenu();
            AllowDrop = true;
            PreviewDragOver += OnDwgDragOver;
            Drop += OnDwgDrop;

            ImageHost.AllowDrop = true;
            ImageHost.PreviewDragOver += OnDwgDragOver;
            ImageHost.Drop += OnDwgDrop;

            // перерисовываем крестики при смене масштаба/плотности/границ
            Vm.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(CoordinateExportingViewModel.EffectiveScale) ||
                    e.PropertyName == nameof(CoordinateExportingViewModel.PixelsPerUnit) ||
                    e.PropertyName == nameof(CoordinateExportingViewModel.MinX) ||
                    e.PropertyName == nameof(CoordinateExportingViewModel.MaxY))
                {
                    RedrawSelectedPoints();
                }
            };

            Focusable = true;
            Focus();
        }


        private DataTemplate CreateLayerItemTemplate()
        {
            var f = new FrameworkElementFactory(typeof(CheckBox));
            f.SetBinding(CheckBox.IsCheckedProperty, new Binding("IsVisible") { Mode = BindingMode.TwoWay });
            f.SetBinding(CheckBox.ContentProperty, new Binding("Name"));
            f.SetValue(FrameworkElement.MarginProperty, new Thickness(6, 2, 6, 2));
            return new DataTemplate { VisualTree = f };
        }

        private ContextMenu BuildLayersMenu()
        {
            var menu = new ContextMenu { StaysOpen = true };

            var items = new ItemsControl
            {
                ItemsSource = Vm.Layers,
                ItemTemplate = CreateLayerItemTemplate(),
                Margin = new Thickness(4)
            };

            var scroll = new ScrollViewer
            {
                MaxHeight = 320,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = items
            };

            var btnAll = new Button { Content = "Все", Margin = new Thickness(0, 0, 6, 0) };
            var btnNone = new Button { Content = "Ничего" };
            btnAll.Click += (_, __) => { foreach (var l in Vm.Layers) l.IsVisible = true; };
            btnNone.Click += (_, __) => { foreach (var l in Vm.Layers) l.IsVisible = false; };

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(6)
            };
            buttons.Children.Add(btnAll);
            buttons.Children.Add(btnNone);

            var panel = new StackPanel();
            panel.Children.Add(scroll);
            panel.Children.Add(new Separator());
            panel.Children.Add(buttons);

            var container = new MenuItem { StaysOpenOnClick = true };
            container.Header = panel;
            menu.Items.Add(container);

            return menu;
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e) => Vm.OpenDwgCommand.Execute(null);
        private void BtnExport_Click(object sender, RoutedEventArgs e) => Vm.ExportCoordsCmd.Execute(null);
        private void BtnSend_Click(object sender, RoutedEventArgs e) => Vm.SendToDataCmd.Execute(null);

        private void BtnRuler_Click(object sender, RoutedEventArgs e)
        {
            Vm.IsMeasureMode = !Vm.IsMeasureMode;
            if (Vm.IsMeasureMode)
            {
                Vm.IsSelectMode = false;
                _measureStart = null;
                _ruler.Visibility = _label.Visibility = Visibility.Collapsed;
            }
        }

        private void BtnSelect_Click(object sender, RoutedEventArgs e)
        {
            Vm.IsSelectMode = !Vm.IsSelectMode;
            if (Vm.IsSelectMode)
            {
                Vm.IsMeasureMode = false;
                SelectedPointsClearOverlay();
            }
        }

        private void BtnLayers_Click(object sender, RoutedEventArgs e)
        {
            if (BtnLayers.ContextMenu == null) return;
            BtnLayers.ContextMenu.DataContext = DataContext;
            BtnLayers.ContextMenu.PlacementTarget = BtnLayers;
            BtnLayers.ContextMenu.IsOpen = true;
        }

        private void SelectedPointsClearOverlay()
        {
            OverlayCanvas.Children.Clear();
            OverlayCanvas.Children.Add(_ruler);
            OverlayCanvas.Children.Add(_label);
            Vm.SelectedPoints.Clear();
            _crossStack.Clear();
        }

        private void Fit() => Vm.FitToViewport(Viewport.ActualWidth, Viewport.ActualHeight);

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
                Vm.SelectedPoints.Add(new System.Windows.Point(cad.X, cad.Y));

                // ВАЖНО: Canvas внутри того же Grid, который масштабируется LayoutTransform,
                // поэтому здесь НЕ умножаем на EffectiveScale
                var px = (cad.X - Vm.MinX) * Vm.PixelsPerUnit;
                var py = (Vm.MaxY - cad.Y) * Vm.PixelsPerUnit;
                DrawCross(px, py);
            }
        }

        private void RedrawSelectedPoints()
        {
            // сохраним линейку и метку
            var keep = new HashSet<UIElement> { _ruler, _label };
            var toRemove = new List<UIElement>();
            foreach (UIElement el in OverlayCanvas.Children)
                if (!keep.Contains(el)) toRemove.Add(el);
            foreach (var el in toRemove) OverlayCanvas.Children.Remove(el);
            _crossStack.Clear();

            foreach (var pt in Vm.SelectedPoints)
            {
                var px = (pt.X - Vm.MinX) * Vm.PixelsPerUnit;
                var py = (Vm.MaxY - pt.Y) * Vm.PixelsPerUnit;
                DrawCross(px, py);
            }
        }


        private void UndoLastPoint()
        {
            if (Vm.SelectedPoints.Count > 0)
            {
                Vm.SelectedPoints.RemoveAt(Vm.SelectedPoints.Count - 1);
                if (_crossStack.Count > 0)
                {
                    var last = _crossStack.Pop();
                    foreach (var el in last)
                        OverlayCanvas.Children.Remove(el);
                }
            }
        }

        private void DrawCross(double x, double y)
        {
            const double sz = 6;
            var l1 = new Line { X1 = x - sz, Y1 = y - sz, X2 = x + sz, Y2 = y + sz, Stroke = Brushes.Red, StrokeThickness = 2 };
            var l2 = new Line { X1 = x - sz, Y1 = y + sz, X2 = x + sz, Y2 = y - sz, Stroke = Brushes.Red, StrokeThickness = 2 };
            OverlayCanvas.Children.Add(l1);
            OverlayCanvas.Children.Add(l2);
            _crossStack.Push(new UIElement[] { l1, l2 });
        }

        private void UserControl_KeyDown(object sender, KeyEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control && e.Key == Key.Z)
            {
                UndoLastPoint();
                e.Handled = true;
                return;
            }

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

        private (double X, double Y) PxToCad(System.Windows.Point p)
        {
            // учитываем зум (LayoutTransform масштабирует грид)
            double s = Math.Max(Vm.EffectiveScale, 1e-9);
            double dwgX = Vm.MinX + p.X / (Vm.PixelsPerUnit * s);
            double dwgY = Vm.MaxY - p.Y / (Vm.PixelsPerUnit * s);
            return (dwgX, dwgY);
        }


        private static bool IsAccepted(string path)
        {
            var ext = System.IO.Path.GetExtension(path);
            return string.Equals(ext, ".dwg", StringComparison.OrdinalIgnoreCase);
        }

        private void OnDwgDragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                bool ok = false;
                foreach (var f in files)
                    if (IsAccepted(f)) { ok = true; break; }
                e.Effects = ok ? DragDropEffects.Copy : DragDropEffects.None;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void OnDwgDrop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);

            string? first = null;
            foreach (var f in files)
                if (IsAccepted(f)) { first = f; break; }
            if (first == null) return;

            Vm.OpenDwgFromPath(first);
        }

        private void GridStepPlus_Click(object sender, RoutedEventArgs e)
        {
            Vm.GridStep = Math.Round(Math.Max(0.1, Vm.GridStep + 1.0), 3);
        }
        private void GridStepMinus_Click(object sender, RoutedEventArgs e)
        {
            Vm.GridStep = Math.Round(Math.Max(0.1, Vm.GridStep - 1.0), 3);
        }
    }
}
