using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;

namespace WinUx.Controls
{
    public partial class Dial : UserControl
    {
        public static readonly DependencyProperty ValueProperty =
            DependencyProperty.Register(nameof(Value), typeof(double), typeof(Dial),
                new FrameworkPropertyMetadata(0.0, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnValueChanged));

        public static readonly DependencyProperty MinimumProperty =
            DependencyProperty.Register(nameof(Minimum), typeof(double), typeof(Dial),
                new PropertyMetadata(0.0, OnRangeChanged));

        public static readonly DependencyProperty MaximumProperty =
            DependencyProperty.Register(nameof(Maximum), typeof(double), typeof(Dial),
                new PropertyMetadata(100.0, OnRangeChanged));

        public static readonly DependencyProperty NotchesVisibleProperty =
            DependencyProperty.Register(nameof(NotchesVisible), typeof(bool), typeof(Dial),
                new PropertyMetadata(true, OnNotchesVisibleChanged));

        public static readonly DependencyProperty NotchTargetProperty =
            DependencyProperty.Register(nameof(NotchTarget), typeof(double), typeof(Dial),
                new PropertyMetadata(10.0, OnNotchTargetChanged));

        public static readonly DependencyProperty DecimalPlacesProperty =
            DependencyProperty.Register(nameof(DecimalPlaces), typeof(int), typeof(Dial),
                new PropertyMetadata(0, OnDecimalPlacesChanged));

        public static readonly DependencyProperty WrappingProperty =
            DependencyProperty.Register(nameof(Wrapping), typeof(bool), typeof(Dial),
                new PropertyMetadata(false));

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

        public bool NotchesVisible
        {
            get => (bool)GetValue(NotchesVisibleProperty);
            set => SetValue(NotchesVisibleProperty, value);
        }

        public double NotchTarget
        {
            get => (double)GetValue(NotchTargetProperty);
            set => SetValue(NotchTargetProperty, value);
        }

        public int DecimalPlaces
        {
            get => (int)GetValue(DecimalPlacesProperty);
            set => SetValue(DecimalPlacesProperty, value);
        }

        public bool Wrapping
        {
            get => (bool)GetValue(WrappingProperty);
            set => SetValue(WrappingProperty, value);
        }

        private bool _isDragging;
        private Point _lastMousePosition;
        private const double StartAngle = 30;  // Start at 30 degrees (bottom-left)
        private const double EndAngle = 330;   // End at 330 degrees (bottom-right)
        private const double TotalAngleRange = 300; // 300 degrees of rotation

        public Dial()
        {
            InitializeComponent();
            Loaded += Dial_Loaded;
        }

        private void Dial_Loaded(object sender, RoutedEventArgs e)
        {
            DrawNotches();
            UpdateVisuals();
        }

        private static void OnValueChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is Dial dial)
            {
                dial.CoerceValue();
                dial.UpdateVisuals();
            }
        }

        private static void OnRangeChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is Dial dial)
            {
                dial.CoerceValue();
                dial.DrawNotches();
                dial.UpdateVisuals();
            }
        }

        private static void OnNotchesVisibleChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is Dial dial)
            {
                dial.DrawNotches();
            }
        }

        private static void OnNotchTargetChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is Dial dial)
            {
                dial.DrawNotches();
            }
        }

        private static void OnDecimalPlacesChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is Dial dial)
            {
                dial.UpdateValueText();
            }
        }

        private void CoerceValue()
        {
            if (Value < Minimum) Value = Minimum;
            if (Value > Maximum) Value = Maximum;
        }

        private void UpdateVisuals()
        {
            UpdateKnobRotation();
            UpdateValueArc();
            UpdateValueText();
        }

        private void UpdateKnobRotation()
        {
            double normalizedValue = (Value - Minimum) / (Maximum - Minimum);
            double angle = StartAngle + (normalizedValue * TotalAngleRange);
            KnobRotation.Angle = angle;
        }

        private void UpdateValueArc()
        {
            double normalizedValue = (Value - Minimum) / (Maximum - Minimum);
            double sweepAngle = normalizedValue * TotalAngleRange;

            // Create arc path
            double centerX = 60;
            double centerY = 60;
            double radius = 50;

            double startAngleRad = (StartAngle - 90) * Math.PI / 180;
            double endAngleRad = (StartAngle + sweepAngle - 90) * Math.PI / 180;

            double startX = centerX + radius * Math.Cos(startAngleRad);
            double startY = centerY + radius * Math.Sin(startAngleRad);
            double endX = centerX + radius * Math.Cos(endAngleRad);
            double endY = centerY + radius * Math.Sin(endAngleRad);

            bool largeArc = sweepAngle > 180;

            if (sweepAngle > 0)
            {
                string pathData = $"M {startX},{startY} A {radius},{radius} 0 {(largeArc ? 1 : 0)} 1 {endX},{endY}";
                ValueArc.Data = Geometry.Parse(pathData);
            }
            else
            {
                ValueArc.Data = Geometry.Parse("M 60,10 A 50,50 0 0 1 60,10");
            }
        }

        private void UpdateValueText()
        {
            ValueText.Text = Value.ToString($"F{DecimalPlaces}", CultureInfo.InvariantCulture);
        }

        private void DrawNotches()
        {
            NotchCanvas.Children.Clear();

            if (!NotchesVisible || NotchTarget <= 0)
                return;

            double range = Maximum - Minimum;
            int notchCount = Math.Max(2, (int)Math.Round(range / NotchTarget) + 1);

            for (int i = 0; i < notchCount; i++)
            {
                double normalizedPosition = (double)i / (notchCount - 1);
                double angle = StartAngle + (normalizedPosition * TotalAngleRange);
                DrawNotch(angle, i % 5 == 0);
            }
        }

        private void DrawNotch(double angle, bool isMajor)
        {
            double centerX = 60;
            double centerY = 60;
            double outerRadius = 58;
            double innerRadius = isMajor ? 50 : 53;

            double angleRad = (angle - 90) * Math.PI / 180;
            double x1 = centerX + outerRadius * Math.Cos(angleRad);
            double y1 = centerY + outerRadius * Math.Sin(angleRad);
            double x2 = centerX + innerRadius * Math.Cos(angleRad);
            double y2 = centerY + innerRadius * Math.Sin(angleRad);

            Line notch = new Line
            {
                X1 = x1,
                Y1 = y1,
                X2 = x2,
                Y2 = y2,
                Stroke = new SolidColorBrush(Color.FromRgb(153, 153, 153)),
                StrokeThickness = isMajor ? 1.5 : 1,
                StrokeStartLineCap = PenLineCap.Round,
                StrokeEndLineCap = PenLineCap.Round
            };

            NotchCanvas.Children.Add(notch);
        }

        private void Dial_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                _isDragging = true;
                _lastMousePosition = e.GetPosition(this);
                Mouse.Capture((IInputElement)sender);
                e.Handled = true;
            }
        }

        private void Dial_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDragging)
                return;

            Point currentPosition = e.GetPosition(this);
            Point center = new Point(ActualWidth / 2, ActualHeight / 2);

            // Calculate angle from center
            double angle = Math.Atan2(currentPosition.Y - center.Y, currentPosition.X - center.X) * 180 / Math.PI;
            angle += 90; // Adjust so 0 degrees is at top

            if (angle < 0) angle += 360;

            // Map angle to value
            double normalizedAngle;

            if (angle < StartAngle)
            {
                normalizedAngle = 0;
            }
            else if (angle > EndAngle && angle < 360)
            {
                normalizedAngle = 1;
            }
            else
            {
                normalizedAngle = (angle - StartAngle) / TotalAngleRange;
            }

            // Clamp normalized angle
            normalizedAngle = Math.Max(0, Math.Min(1, normalizedAngle));

            // Update value
            double newValue = Minimum + (normalizedAngle * (Maximum - Minimum));
            Value = newValue;

            e.Handled = true;
        }

        private void Dial_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left && _isDragging)
            {
                _isDragging = false;
                Mouse.Capture(null);
                e.Handled = true;
            }
        }

        private void Dial_MouseLeave(object sender, MouseEventArgs e)
        {
            if (_isDragging)
            {
                _isDragging = false;
                Mouse.Capture(null);
            }
        }

        protected override void OnMouseWheel(MouseWheelEventArgs e)
        {
            base.OnMouseWheel(e);

            double increment = (Maximum - Minimum) / 100.0;
            if (e.Delta > 0)
            {
                Value = Math.Min(Value + increment, Maximum);
            }
            else
            {
                Value = Math.Max(Value - increment, Minimum);
            }

            e.Handled = true;
        }
    }
}