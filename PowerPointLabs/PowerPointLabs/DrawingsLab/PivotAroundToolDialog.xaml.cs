using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using PowerPointLabs.DataSources;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.DrawingsLab
{
    /// <summary>
    /// Interaction logic for AlignmentDialogVertical.xaml
    /// </summary>
    public partial class PivotAroundToolDialog : Window
    {
        #region Properties
        public double StartAngle
        {
            get { return dataSource.StartAngle; }
        }

        public double AngleDifference
        {
            get { return dataSource.AngleDifference; }
        }

        public int Copies
        {
            get
            {
                if (dataSource.Copies < 1) return 1;
                return dataSource.Copies;
            }
        }

        public bool IsExtend
        {
            get { return dataSource.IsExtend; }
        }

        public bool RotateShape
        {
            get { return dataSource.RotateShape; }
        }

        public bool FixOriginalLocation
        {
            get { return dataSource.FixOriginalLocation; }
        }

        public DrawingsLabPivotAroundDataSource.Alignment PivotAnchor
        {
            get { return dataSource.PivotAnchor; }
        }

        public double PivotAnchorFractionX
        {
            get { return ToAnchorFractionX(dataSource.PivotAnchor); }
        }

        public double PivotAnchorFractionY
        {
            get { return ToAnchorFractionY(dataSource.PivotAnchor); }
        }

        public double PivotCenterX
        {
            get
            {
                double p = PivotAnchorFractionX;
                return _pivotLeft * (1 - p) + _pivotRight * p;
            }
        }

        public double PivotCenterY
        {
            get
            {
                double p = PivotAnchorFractionY;
                return _pivotTop * (1 - p) + _pivotBottom * p;
            }
        }

        public double SourceAnchorFractionX
        {
            get { return ToAnchorFractionX(dataSource.SourceAnchor); }
        }

        public double SourceAnchorFractionY
        {
            get { return ToAnchorFractionY(dataSource.SourceAnchor); }
        }

        public double SourceCenterX
        {
            get
            {
                double p = SourceAnchorFractionX;
                return _sourceLeft * (1 - p) + _sourceRight * p;
            }
        }

        public double SourceCenterY
        {
            get
            {
                double p = SourceAnchorFractionY;
                return _sourceTop * (1 - p) + _sourceBottom * p;
            }
        }

        private static double ToAnchorFractionX(DrawingsLabPivotAroundDataSource.Alignment anchor)
        {
            switch (anchor)
            {
                case DrawingsLabPivotAroundDataSource.Alignment.TopLeft:
                case DrawingsLabPivotAroundDataSource.Alignment.MiddleLeft:
                case DrawingsLabPivotAroundDataSource.Alignment.BottomLeft:
                    return 0;
                case DrawingsLabPivotAroundDataSource.Alignment.TopCenter:
                case DrawingsLabPivotAroundDataSource.Alignment.MiddleCenter:
                case DrawingsLabPivotAroundDataSource.Alignment.BottomCenter:
                    return 0.5;
                case DrawingsLabPivotAroundDataSource.Alignment.TopRight:
                case DrawingsLabPivotAroundDataSource.Alignment.MiddleRight:
                case DrawingsLabPivotAroundDataSource.Alignment.BottomRight:
                    return 1;
            }
            return 0.5;
        }

        private static double ToAnchorFractionY(DrawingsLabPivotAroundDataSource.Alignment anchor)
        {
            switch (anchor)
            {
                case DrawingsLabPivotAroundDataSource.Alignment.TopLeft:
                case DrawingsLabPivotAroundDataSource.Alignment.TopCenter:
                case DrawingsLabPivotAroundDataSource.Alignment.TopRight:
                    return 0;
                case DrawingsLabPivotAroundDataSource.Alignment.MiddleLeft:
                case DrawingsLabPivotAroundDataSource.Alignment.MiddleCenter:
                case DrawingsLabPivotAroundDataSource.Alignment.MiddleRight:
                    return 0.5;
                case DrawingsLabPivotAroundDataSource.Alignment.BottomLeft:
                case DrawingsLabPivotAroundDataSource.Alignment.BottomCenter:
                case DrawingsLabPivotAroundDataSource.Alignment.BottomRight:
                    return 1;
            }
            return 0.5;
        }
        #endregion

        private DrawingsLabPivotAroundDataSource dataSource;
        private readonly double _sourceLeft;
        private readonly double _sourceTop;
        private readonly double _sourceRight;
        private readonly double _sourceBottom;
        private readonly double _sourceRotation;
        private readonly double _pivotLeft;
        private readonly double _pivotTop;
        private readonly double _pivotRight;
        private readonly double _pivotBottom;

        private double _baseRotation;

        private const double MARGIN = 5.0;

        private double CanvasWidth
        {
            get { return Math.Max(0, PivotCanvas.ActualWidth - 2*MARGIN); }
        }

        private double CanvasHeight
        {
            get { return Math.Max(0, PivotCanvas.ActualHeight - 2*MARGIN); }
        }

        public PivotAroundToolDialog(Shape sourceShape, Shape pivotShape)
        {
            _sourceLeft = sourceShape.Left;
            _sourceTop = sourceShape.Top;
            _sourceRight = sourceShape.Left + sourceShape.Width;
            _sourceBottom = sourceShape.Top + sourceShape.Height;
            _sourceRotation = sourceShape.Rotation;

            _pivotLeft = pivotShape.Left;
            _pivotTop = pivotShape.Top;
            _pivotRight = pivotShape.Left + pivotShape.Width;
            _pivotBottom = pivotShape.Top + pivotShape.Height;

            InitializeComponent();
            InitialiseDataSource();
            SetDefaultAngle();

            dataSource.PropertyChanged += OnPropertyChanged;
        }

        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "PivotAnchor" || e.PropertyName == "SourceAnchor" || e.PropertyName == "FixOriginalLocation")
            {
                if (dataSource.FixOriginalLocation)
                {
                    SetDefaultAngle();
                }
            }
            DrawAlignmentCanvas();
        }

        private void SetDefaultAngle()
        {
            var angle = Math.Atan2(SourceCenterY - PivotCenterY, SourceCenterX - PivotCenterX)*180/Math.PI;
            dataSource.StartAngle = angle;
            _baseRotation = _sourceRotation - angle;
        }

        private double ToActualX(double f)
        {
            if (CanvasWidth > CanvasHeight)
            {
                return MARGIN + (CanvasWidth - CanvasHeight)/2 + CanvasHeight*f;
            }
            else
            {
                return MARGIN + CanvasWidth*f;
            }
        }

        private double ToActualY(double f)
        {
            if (CanvasHeight > CanvasWidth)
            {
                return MARGIN + (CanvasHeight - CanvasWidth)/2 + CanvasWidth*f;
            }
            else
            {
                return MARGIN + CanvasHeight * f;
            }
        }

        private double ToActualLength(double f)
        {
            return Math.Min(CanvasHeight, CanvasWidth)*f;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            DrawAlignmentCanvas();
        }

        private void DrawAlignmentCanvas()
        {
            PivotCanvas.Children.Clear();
            if (Copies > 1000) return; // Don't draw anything if too many rectangles.

            if (Copies > 1)
            {
                DrawOverlappingArc(ToActualX(0.5), ToActualY(0.5), ToActualLength(0.1), ToActualLength(0.4), StartAngle, AngleDifference);
            }

            double angleStep = AngleDifference;
            if (Copies > 1 && !IsExtend)
            {
                angleStep /= (Copies - 1);
            }

            double circleRadius = ToActualLength(0.02);
            double midpointX = ToActualX(0.5);
            double midpointY = ToActualY(0.5);
            double rectWidth = ToActualLength(0.18);
            double rectHeight = ToActualLength(0.15);
            double pivotRectWidth = ToActualLength(0.12);
            double pivotRectHeight = ToActualLength(0.11);

            DrawRotatedRect(midpointX, midpointY, PivotAnchorFractionX, PivotAnchorFractionY,
                            pivotRectWidth, pivotRectHeight, 0, Brushes.SlateGray);

            for (int i = 0; i < Copies; ++i)
            {
                double angle = (StartAngle + angleStep*i);
                double angleRad = angle*Math.PI/180;
                double cx = ToActualX(0.5 + Math.Cos(angleRad)*0.4);
                double cy = ToActualY(0.5 + Math.Sin(angleRad)*0.4);
                double rotation = RotateShape ? angle + _baseRotation : StartAngle + _baseRotation;

                var rectColour = Brushes.Green;
                var lineColour = Brushes.CornflowerBlue;
                var circleColour = Brushes.DodgerBlue;

                if (i == 0)
                {
                    rectColour = Brushes.LimeGreen;
                    lineColour = Brushes.PaleVioletRed;
                    circleColour = Brushes.MediumVioletRed;
                }

                DrawRotatedRect(cx, cy, SourceAnchorFractionX, SourceAnchorFractionY, rectWidth, rectHeight, rotation, rectColour);
                DrawLine(midpointX, midpointY, cx, cy, lineColour);
                DrawCircle(cx, cy, circleRadius, circleColour);
            }
        }


        private void DrawOverlappingArc(double cx, double cy, double innerRadius, double outerRadius, double angleStart,
            double angleDifference)
        {
            Brush colour = null;
            if (angleDifference > 0)
            {
                // Clockwise
                colour = Brushes.Orange;
            }
            else
            {
                // Anticlockwise
                colour = Brushes.DodgerBlue;
                angleStart += angleDifference;
                angleDifference = -angleDifference;
            }

            while (angleDifference >= 360)
            {
                DrawArc(cx, cy, innerRadius, outerRadius, angleStart, angleStart + 359.9999, colour);
                angleDifference -= 360;
            }
            DrawArc(cx, cy, innerRadius, outerRadius, angleStart, angleStart + angleDifference, colour);
        }

        private void DrawArc(double cx, double cy, double innerRadius, double outerRadius, double angleStart,
            double angleEnd, Brush colour)
        {
            // Assumption: angleStart < angleEnd.
            angleStart *= Math.PI / 180;
            angleEnd *= Math.PI / 180;

            var path = new Path();
            Canvas.SetLeft(path, cx);
            Canvas.SetTop(path, cy);
            path.Fill = colour;
            path.Opacity = 0.2f;

            var pathGeometry = new PathGeometry();
            var innerStartPoint = new Point(Math.Cos(angleStart) * innerRadius, Math.Sin(angleStart) * innerRadius);
            var innerEndPoint = new Point(Math.Cos(angleEnd) * innerRadius, Math.Sin(angleEnd) * innerRadius);
            var outerStartPoint = new Point(Math.Cos(angleStart) * outerRadius, Math.Sin(angleStart) * outerRadius);
            var outerEndPoint = new Point(Math.Cos(angleEnd) * outerRadius, Math.Sin(angleEnd) * outerRadius);

            var pathFigure = new PathFigure
            {
                StartPoint = innerStartPoint,
                IsClosed = true
            };
            var seg1 = new LineSegment(outerStartPoint, true);
            var seg2 = new ArcSegment
            {
                IsLargeArc = Math.Abs(angleEnd - angleStart) >= Math.PI,
                Point = outerEndPoint,
                Size = new Size(outerRadius, outerRadius),
                SweepDirection = SweepDirection.Clockwise
            };
            var seg3 = new LineSegment(innerEndPoint, true);
            var seg4 = new ArcSegment
            {
                IsLargeArc = Math.Abs(angleEnd - angleStart) >= Math.PI,
                Point = innerStartPoint,
                Size = new Size(innerRadius, innerRadius),
                SweepDirection = SweepDirection.Counterclockwise
            };

            pathFigure.Segments.Add(seg1);
            pathFigure.Segments.Add(seg2);
            pathFigure.Segments.Add(seg3);
            pathFigure.Segments.Add(seg4);
            pathGeometry.Figures.Add(pathFigure);
            path.Data = pathGeometry;
            PivotCanvas.Children.Add(path);
        }

        private void DrawLine(double x1, double y1, double x2, double y2, Brush colour)
        {
            var line = new Line
            {
                X1 = x1,
                Y1 = y1,
                X2 = x2,
                Y2 = y2,
                Stroke = colour,
                StrokeThickness = 2
            };
            PivotCanvas.Children.Add(line);
        }

        private void DrawRect(double x, double y, double width, double height, Brush colour)
        {
            var rect = new Rectangle
            {
                Width = width,
                Height = height,
                Stroke = Brushes.LightGray,
                Fill = colour,
                StrokeThickness = 2
            };
            Canvas.SetLeft(rect, x);
            Canvas.SetTop(rect, y);
            PivotCanvas.Children.Add(rect);
        }

        private void DrawCircle(double cx, double cy, double radius, Brush colour)
        {
            var ellipse = new Ellipse
            {
                Width = radius*2,
                Height = radius*2,
                Fill = colour
            };
            Canvas.SetLeft(ellipse, cx-radius);
            Canvas.SetTop(ellipse, cy-radius);
            PivotCanvas.Children.Add(ellipse);
        }

        private void DrawRotatedRect(double cx, double cy, double anchorX, double anchorY, double width, double height, double angle, Brush colour)
        {
            var rect = new Rectangle
            {
                Width = width,
                Height = height,
                Stroke = Brushes.Black,
                Fill = colour,
                StrokeThickness = 1,
                Opacity = 0.7,
                RenderTransform = new RotateTransform(angle, width*anchorX, height*anchorY)
            };
            Canvas.SetLeft(rect, cx - width*anchorX);
            Canvas.SetTop(rect, cy - height*anchorY);
            PivotCanvas.Children.Add(rect);
        }

        private void InitialiseDataSource()
        {
            dataSource = FindResource("DataSource") as DrawingsLabPivotAroundDataSource;
        }

        private void ButtomDialogOk_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }
    }
}
