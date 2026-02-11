using CodeReportTracker.Core.Models;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace CodeReportTracker.Components
{
    public partial class CodeItemEditDialog : Window
    {
        private readonly CodeItem _original;
        private readonly CodeItem _working;

        public CodeItemEditDialog(CodeItem item)
        {
            InitializeComponent();

            _original = item ?? throw new ArgumentNullException(nameof(item));
            _working = new CodeItem();

            // Copy properties from original into working copy for editing
            CopyAllProperties(_original, _working);

            DataContext = _working;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Copy changed properties back to original
                CopyAllProperties(_working, _original);

                // If host uses per-row EditAction/command, keep that behavior as well.
                _original.EditAction?.Invoke(_original);

                DialogResult = true;
            }
            catch
            {
                DialogResult = false;
            }
            finally
            {
                Close();
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        // Shallow copy of public instance properties with public getter/setter.
        private static void CopyAllProperties(CodeItem source, CodeItem target)
        {
            if (source == null || target == null) return;

            var type = typeof(CodeItem);
            var props = type.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);

            foreach (var p in props)
            {
                if (!p.CanRead || !p.CanWrite) continue;
                // Skip actions/commands to avoid overwriting host wiring
                if (p.PropertyType == typeof(Action<string>) || p.PropertyType == typeof(Action<CodeItem>) || typeof(System.Windows.Input.ICommand).IsAssignableFrom(p.PropertyType))
                    continue;

                try
                {
                    var value = p.GetValue(source);
                    p.SetValue(target, value);
                }
                catch
                {
                    // ignore individual property copy failures
                }
            }
        }

        // ---- Auto-expand description TextBox helpers ----

        private void DescriptionTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            AdjustDescriptionHeight();
        }

        private void DescriptionTextBox_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            // Recalculate when width changes so wrapping affects height
            AdjustDescriptionHeight();
        }

        private void DescriptionTextBox_Loaded(object sender, RoutedEventArgs e)
        {
            // Defer sizing until layout and ActualWidth are available
            Dispatcher.BeginInvoke(new Action(AdjustDescriptionHeight), System.Windows.Threading.DispatcherPriority.Loaded);
        }

        private void AdjustDescriptionHeight()
        {
            if (DescriptionTextBox == null) return;

            // Create a TextBlock to measure wrapped height using the TextBox width
            var tb = DescriptionTextBox;

            // If ActualWidth is not available yet, skip (will be retried on SizeChanged)
            if (tb.ActualWidth <= 0) return;

            var measure = new TextBlock
            {
                Text = string.IsNullOrEmpty(tb.Text) ? " " : tb.Text + " ",
                TextWrapping = TextWrapping.Wrap,
                FontFamily = tb.FontFamily,
                FontSize = tb.FontSize,
                FontStretch = tb.FontStretch,
                FontStyle = tb.FontStyle,
                FontWeight = tb.FontWeight
            };

            // approximate available width for text inside the TextBox (subtract padding and a small margin)
            var availableWidth = Math.Max(0.0, tb.ActualWidth - tb.Padding.Left - tb.Padding.Right - 4.0);
            measure.Width = availableWidth;

            // Measure with infinite height so TextBlock returns needed height
            measure.Measure(new Size(availableWidth, double.PositiveInfinity));
            var desired = measure.DesiredSize.Height + tb.Padding.Top + tb.Padding.Bottom;

            // clamp to reasonable min/max to avoid uncontrolled growth
            var min = tb.MinHeight > 0 ? tb.MinHeight : 80.0;
            var max = 600.0;

            tb.Height = Math.Min(Math.Max(desired, min), max);
        }
    }
}