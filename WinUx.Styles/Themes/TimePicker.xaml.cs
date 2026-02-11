using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace WinUx.Controls
{
    public partial class TimePicker : UserControl
    {
        public TimePicker()
        {
            InitializeComponent();
            Loaded += OnLoaded;
            // Subscribe to the IsEnabledChanged event instead of attempting to override a non-existent method
            IsEnabledChanged += OnIsEnabledChangedHandler;
        }

        private void OnLoaded(object? sender, RoutedEventArgs e)
        {
            PART_TextBox.LostFocus += TextBox_LostFocus;
            PART_TextBox.KeyDown += TextBox_KeyDown;
            PART_UpButton.Click += UpButton_Click;
            PART_DownButton.Click += DownButton_Click;

            UpdateTextFromValue();
            UpdateEnabledState();
        }

        #region Dependency Properties

        public static readonly DependencyProperty ValueProperty =
            DependencyProperty.Register(
                nameof(Value),
                typeof(DateTime?),
                typeof(TimePicker),
                new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnValueChanged));

        public DateTime? Value
        {
            get => (DateTime?)GetValue(ValueProperty);
            set => SetValue(ValueProperty, value);
        }

        private static void OnValueChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is TimePicker tp)
                tp.UpdateTextFromValue();
        }

        public static readonly DependencyProperty FormatStringProperty =
            DependencyProperty.Register(
                nameof(FormatString),
                typeof(string),
                typeof(TimePicker),
                new PropertyMetadata("HH:mm"));

        public string FormatString
        {
            get => (string)GetValue(FormatStringProperty);
            set => SetValue(FormatStringProperty, value);
        }

        public static readonly DependencyProperty MinuteStepProperty =
            DependencyProperty.Register(
                nameof(MinuteStep),
                typeof(int),
                typeof(TimePicker),
                new PropertyMetadata(1));

        public int MinuteStep
        {
            get => (int)GetValue(MinuteStepProperty);
            set => SetValue(MinuteStepProperty, value);
        }

        #endregion

        private void UpdateTextFromValue()
        {
            if (PART_TextBox == null)
                return;

            if (Value.HasValue)
            {
                try
                {
                    PART_TextBox.Text = Value.Value.ToString(FormatString ?? "HH:mm", CultureInfo.CurrentCulture);
                }
                catch
                {
                    PART_TextBox.Text = Value.Value.ToString(CultureInfo.CurrentCulture);
                }
            }
            else
            {
                PART_TextBox.Text = string.Empty;
            }
        }

        private void TextBox_LostFocus(object? sender, RoutedEventArgs e) => CommitText();

        private void TextBox_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                CommitText();
                e.Handled = true;
            }
        }

        private void CommitText()
        {
            if (PART_TextBox == null)
                return;

            var txt = PART_TextBox.Text?.Trim();
            if (string.IsNullOrEmpty(txt))
            {
                Value = null;
                return;
            }

            if (!string.IsNullOrEmpty(FormatString) &&
                DateTime.TryParseExact(txt, FormatString, CultureInfo.CurrentCulture, DateTimeStyles.None, out var exact))
            {
                Value = exact;
                return;
            }

            if (DateTime.TryParse(txt, CultureInfo.CurrentCulture, DateTimeStyles.None, out var parsed))
            {
                Value = parsed;
                return;
            }

            // revert on parse failure
            UpdateTextFromValue();
        }

        private void UpButton_Click(object? sender, RoutedEventArgs e) => StepMinutes(Math.Abs(MinuteStep));
        private void DownButton_Click(object? sender, RoutedEventArgs e) => StepMinutes(-Math.Abs(MinuteStep));

        private void StepMinutes(int minutes)
        {
            var baseTime = Value ?? DateTime.Now;
            Value = baseTime.AddMinutes(minutes);
        }

        private void UpdateEnabledState()
        {
            if (PART_TextBox != null)
                PART_TextBox.IsEnabled = IsEnabled;

            if (PART_UpButton != null)
                PART_UpButton.IsEnabled = IsEnabled;

            if (PART_DownButton != null)
                PART_DownButton.IsEnabled = IsEnabled;
        }

        // Event handler for IsEnabledChanged - replaces the invalid override
        private void OnIsEnabledChangedHandler(object? sender, DependencyPropertyChangedEventArgs e)
        {
            UpdateEnabledState();
        }
    }
}