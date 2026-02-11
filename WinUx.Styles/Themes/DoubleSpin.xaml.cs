using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace WinUx.Controls
{
    public partial class DoubleSpin : UserControl
    {
        public static readonly DependencyProperty ValueProperty =
            DependencyProperty.Register(nameof(Value), typeof(double), typeof(DoubleSpin),
                new FrameworkPropertyMetadata(0.0, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnValueChanged));

        public static readonly DependencyProperty MinimumProperty =
            DependencyProperty.Register(nameof(Minimum), typeof(double), typeof(DoubleSpin),
                new PropertyMetadata(double.MinValue));

        public static readonly DependencyProperty MaximumProperty =
            DependencyProperty.Register(nameof(Maximum), typeof(double), typeof(DoubleSpin),
                new PropertyMetadata(double.MaxValue));

        public static readonly DependencyProperty IncrementProperty =
            DependencyProperty.Register(nameof(Increment), typeof(double), typeof(DoubleSpin),
                new PropertyMetadata(1.0));

        public static readonly DependencyProperty DecimalPlacesProperty =
            DependencyProperty.Register(nameof(DecimalPlaces), typeof(int), typeof(DoubleSpin),
                new PropertyMetadata(0, OnDecimalPlacesChanged));

        public double Value
        {
            get => (double)GetValue(ValueProperty);
            set => SetValue(ValueProperty, value);
        }

        public double Minimum
        {
            get => (double)GetValue(MinimumProperty);
            set => SetValue(MinimumProperty, value);
        }

        public double Maximum
        {
            get => (double)GetValue(MaximumProperty);
            set => SetValue(MaximumProperty, value);
        }

        public double Increment
        {
            get => (double)GetValue(IncrementProperty);
            set => SetValue(IncrementProperty, value);
        }

        public int DecimalPlaces
        {
            get => (int)GetValue(DecimalPlacesProperty);
            set => SetValue(DecimalPlacesProperty, value);
        }

        private bool _isUpdating;

        public DoubleSpin()
        {
            InitializeComponent();
            UpdateTextBox();
        }

        private static void OnValueChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is DoubleSpin control && !control._isUpdating)
            {
                control.CoerceValue();
                control.UpdateTextBox();
            }
        }

        private static void OnDecimalPlacesChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is DoubleSpin control)
            {
                control.UpdateTextBox();
            }
        }

        private void CoerceValue()
        {
            if (Value < Minimum) Value = Minimum;
            if (Value > Maximum) Value = Maximum;
        }

        private void UpdateTextBox()
        {
            _isUpdating = true;
            ValueTextBox.Text = Value.ToString($"F{DecimalPlaces}", CultureInfo.InvariantCulture);
            _isUpdating = false;
        }

        private void ValueTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Allow numbers, decimal point, minus sign
            e.Handled = !IsTextAllowed(e.Text);
        }

        private static bool IsTextAllowed(string text)
        {
            return Regex.IsMatch(text, @"^[0-9.\-]+$");
        }

        private void ValueTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdating) return;

            if (double.TryParse(ValueTextBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double newValue))
            {
                _isUpdating = true;
                Value = newValue;
                CoerceValue();
                _isUpdating = false;
            }
        }

        private void ValueTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            UpdateTextBox();
        }

        private void UpButton_Click(object sender, RoutedEventArgs e)
        {
            Value = Math.Min(Value + Increment, Maximum);
        }

        private void DownButton_Click(object sender, RoutedEventArgs e)
        {
            Value = Math.Max(Value - Increment, Minimum);
        }
    }
}