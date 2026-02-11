using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;

namespace WinUx.Controls
{
    public partial class RibbonMenu : UserControl
    {
        private bool _hostHandlersAttached;

        public RibbonMenu()
        {
            InitializeComponent();

            // Use the control itself as the DataContext for internal bindings
            DataContext = this;

            // Wire events to keep toggle and popup in sync
            PART_Toggle.Checked += OnToggleChecked;
            PART_Toggle.Unchecked += OnToggleUnchecked;
            PART_Popup.Closed += PART_Popup_Closed;

            Loaded += RibbonMenu_Loaded;
        }

        private void RibbonMenu_Loaded(object sender, RoutedEventArgs e)
        {
            if (ItemTemplate == null && Resources.Contains("DefaultRibbonMenuItemTemplate"))
            {
                ItemTemplate = (DataTemplate)Resources["DefaultRibbonMenuItemTemplate"];
            }
        }

        // Title shown on the toggle button
        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register(nameof(Title), typeof(string), typeof(RibbonMenu), new PropertyMetadata("Menu"));

        public string Title
        {
            get => (string)GetValue(TitleProperty);
            set => SetValue(TitleProperty, value);
        }

        public static readonly DependencyProperty ItemsSourceProperty =
            DependencyProperty.Register(nameof(ItemsSource), typeof(IEnumerable), typeof(RibbonMenu), new PropertyMetadata(null));

        public IEnumerable ItemsSource
        {
            get => (IEnumerable)GetValue(ItemsSourceProperty);
            set => SetValue(ItemsSourceProperty, value);
        }

        public static readonly DependencyProperty ItemCommandProperty =
            DependencyProperty.Register(nameof(ItemCommand), typeof(ICommand), typeof(RibbonMenu), new PropertyMetadata(null));

        public ICommand ItemCommand
        {
            get => (ICommand)GetValue(ItemCommandProperty);
            set => SetValue(ItemCommandProperty, value);
        }

        public static readonly DependencyProperty ItemTemplateProperty =
            DependencyProperty.Register(nameof(ItemTemplate), typeof(DataTemplate), typeof(RibbonMenu), new PropertyMetadata(null));

        public DataTemplate ItemTemplate
        {
            get => (DataTemplate)GetValue(ItemTemplateProperty);
            set => SetValue(ItemTemplateProperty, value);
        }

        public static readonly DependencyProperty TabsSourceProperty =
            DependencyProperty.Register(nameof(TabsSource), typeof(IEnumerable), typeof(RibbonMenu), new PropertyMetadata(null));

        public IEnumerable TabsSource
        {
            get => (IEnumerable)GetValue(TabsSourceProperty);
            set => SetValue(TabsSourceProperty, value);
        }

        private void OnToggleChecked(object sender, RoutedEventArgs e)
        {
            var win = Window.GetWindow(this);
            if (win != null)
            {
                PART_Popup.PlacementTarget = win;
                PART_Popup.Placement = PlacementMode.Relative;
                PART_Popup.HorizontalOffset = 0;
                PART_Popup.VerticalOffset = 0;
                PART_Popup.Width = win.ActualWidth;

                if (!_hostHandlersAttached)
                {
                    win.SizeChanged += HostWindow_SizeChanged;
                    win.LocationChanged += HostWindow_LocationChanged;
                    _hostHandlersAttached = true;
                }
            }

            PART_Popup.IsOpen = true;
        }

        private void HostWindow_SizeChanged(object? sender, SizeChangedEventArgs e)
        {
            if (sender is Window w)
            {
                PART_Popup.Width = e.NewSize.Width;
            }
        }

        private void HostWindow_LocationChanged(object? sender, EventArgs e)
        {
            PART_Popup.HorizontalOffset = 0;
            PART_Popup.VerticalOffset = 0;
        }

        private void OnToggleUnchecked(object sender, RoutedEventArgs e)
        {
            PART_Popup.IsOpen = false;
        }

        private void PART_Popup_Closed(object sender, EventArgs e)
        {
            PART_Toggle.IsChecked = false;

            var win = Window.GetWindow(this);
            if (win != null && _hostHandlersAttached)
            {
                win.SizeChanged -= HostWindow_SizeChanged;
                win.LocationChanged -= HostWindow_LocationChanged;
                _hostHandlersAttached = false;
            }

            PART_Popup.PlacementTarget = PART_Toggle;
            PART_Popup.Placement = PlacementMode.Bottom;
            PART_Popup.Width = double.NaN;
            PART_Popup.HorizontalOffset = 0;
            PART_Popup.VerticalOffset = 0;
        }
    }

    // Public model classes used by XAML DataTemplates.

    public enum RibbonButtonSize
    {
        Small = 0,
        Large = 1
    }

    public class RibbonTab
    {
        public string Header { get; set; } = string.Empty;
        public IEnumerable<RibbonGroup>? Groups { get; set; }
    }

    public class RibbonGroup
    {
        public string Header { get; set; } = string.Empty;
        public IEnumerable? Items { get; set; }
    }

    public class RibbonButtonItem
    {
        public string Label { get; set; } = string.Empty;
        public object? Icon { get; set; }
        public ICommand? Command { get; set; }
        public object? CommandParameter { get; set; }
        public RibbonButtonSize Size { get; set; } = RibbonButtonSize.Small;
    }

    public class RibbonToggleItem
    {
        public string Label { get; set; } = string.Empty;
        public object? Icon { get; set; }
        public bool IsChecked { get; set; }
        public ICommand? Command { get; set; }
        public object? CommandParameter { get; set; }
        public RibbonButtonSize Size { get; set; } = RibbonButtonSize.Small;
    }

    public class RibbonSeparatorItem
    {
    }

    public class RibbonCustomItem
    {
        public object? Content { get; set; }
    }
}