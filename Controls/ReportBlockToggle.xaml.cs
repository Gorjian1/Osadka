// File: Controls/ReportBlockToggle.xaml.cs
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Osadka.ViewModels;

namespace Osadka.Controls
{
    /// <summary>
    /// Встраиваемая «галочка» для включения/отключения блока отчёта.
    /// Ищет блок по ключу в ReportOutputSettings и двунаправленно связывает IsEnabled.
    /// </summary>
    public partial class ReportBlockToggle : UserControl
    {
        public ReportBlockToggle()
        {
            InitializeComponent();
            Loaded += OnLoaded;
        }

        #region Dependency Properties

        /// <summary>Ключ блока (ReportOutputSettings.Blocks[].Key). Если ключ не найден — пробует Title.</summary>
        public string BlockKey
        {
            get => (string)GetValue(BlockKeyProperty);
            set => SetValue(BlockKeyProperty, value);
        }
        public static readonly DependencyProperty BlockKeyProperty =
            DependencyProperty.Register(nameof(BlockKey), typeof(string), typeof(ReportBlockToggle),
                new PropertyMetadata(string.Empty, OnParamsChanged));

        /// <summary>Подпись рядом с чекбоксом. Если не задана — берётся Title блока.</summary>
        public object? Label
        {
            get => GetValue(LabelProperty);
            set => SetValue(LabelProperty, value);
        }
        public static readonly DependencyProperty LabelProperty =
            DependencyProperty.Register(nameof(Label), typeof(object), typeof(ReportBlockToggle),
                new PropertyMetadata(null));

        /// <summary>Явно переданный Settings. Если не задан, контрол попытается найти GeneralReportViewModel вверх по визуальному дереву.</summary>
        public ReportOutputSettings? Settings
        {
            get => (ReportOutputSettings?)GetValue(SettingsProperty);
            set => SetValue(SettingsProperty, value);
        }
        public static readonly DependencyProperty SettingsProperty =
            DependencyProperty.Register(nameof(Settings), typeof(ReportOutputSettings), typeof(ReportBlockToggle),
                new PropertyMetadata(null, OnParamsChanged));

        /// <summary>Показывать ли кнопку-подсказку «i» с перечнем тегов.</summary>
        public bool ShowInfo
        {
            get => (bool)GetValue(ShowInfoProperty);
            set => SetValue(ShowInfoProperty, value);
        }
        public static readonly DependencyProperty ShowInfoProperty =
            DependencyProperty.Register(nameof(ShowInfo), typeof(bool), typeof(ReportBlockToggle),
                new PropertyMetadata(true));

        #endregion

        private void OnLoaded(object? sender, RoutedEventArgs e)
        {
            TryWire();
        }

        private static void OnParamsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ReportBlockToggle t && t.IsLoaded)
                t.TryWire();
        }

        private void TryWire()
        {
            // 1) Где взять Settings:
            var settings = Settings ?? FindSettingsInAncestors(this);
            if (settings == null) return;

            // 2) Найти блок по ключу (или по Title — фоллбэк):
            BlockSetting? block = settings.FindBlock(BlockKey);
            if (block == null)
            {
                // Нет ключа/тайтла — ничего не биндим.
                this.DataContext = null;
                return;
            }

            // 3) Биндим DataContext контролу на BlockSetting
            this.DataContext = block;

            // 4) Если Label не задан — используем Title блока
            if (Label == null)
                SetCurrentValue(LabelProperty, block.Title);
        }

        private static ReportOutputSettings? FindSettingsInAncestors(DependencyObject start)
        {
            DependencyObject? d = start;
            while (d != null)
            {
                if (d is FrameworkElement fe)
                {
                    if (fe.DataContext is GeneralReportViewModel gvm && gvm.Settings != null)
                        return gvm.Settings;
                }
                d = VisualTreeHelper.GetParent(d);
            }
            return null;
        }
    }
}
