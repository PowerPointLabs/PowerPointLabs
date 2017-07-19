using System;
using System.Collections.Generic;
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

namespace PowerPointLabs.DrawingsLab
{
    /// <summary>
    /// Interaction logic for AlignmentDialogVertical.xaml
    /// </summary>
    public partial class MultiCloneGridDialog : Window
    {
        private DrawingsLabMultiCloneGridDataSource dataSource;
        private readonly float _sourceLeft;
        private readonly float _sourceTop;
        private readonly float _targetLeft;
        private readonly float _targetTop;
        
        private const int RATIO_RECT = 3;
        private const int RATIO_GAP = 1;
        private const double MARGIN = 5.0;

        private const int MIN_COPIES = 2;

        public MultiCloneGridDialog(float sourceLeft, float sourceTop, float targetLeft, float targetTop)
        {
            _sourceLeft = sourceLeft;
            _sourceTop = sourceTop;
            _targetLeft = targetLeft;
            _targetTop = targetTop;

            InitializeComponent();
            InitialiseDataSource();

            dataSource.PropertyChanged += (s, e) => DrawAlignmentCanvas();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            DrawAlignmentCanvas();
        }

        private void DrawAlignmentCanvas()
        {
            GridCanvas.Children.Clear();
            if (XCopies * YCopies > 2500)
            {
                return; // Don't draw anything if too many rectangles.
            }

            var drawWidth = GridCanvas.ActualWidth - 2*MARGIN;
            var drawHeight = GridCanvas.ActualHeight - 2*MARGIN;

            int xDivisions = XCopies * RATIO_RECT + (XCopies - 1) * RATIO_GAP;
            int yDivisions = YCopies * RATIO_RECT + (YCopies - 1) * RATIO_GAP;
            
            double gridTileWidth = Math.Min(drawWidth / xDivisions, drawHeight / yDivisions);
            double rectSize = RATIO_RECT*gridTileWidth;
            double intervalSize = (RATIO_RECT+RATIO_GAP)*gridTileWidth;

            int sourceIndexX, sourceIndexY, targetIndexX, targetIndexY;
            ComputeIndexes(out sourceIndexX, out sourceIndexY, out targetIndexX, out targetIndexY);

            for (int y = 0; y < YCopies; ++y)
            {
                for (int x = 0; x < XCopies; ++x)
                {
                    var x1 = MARGIN + x * intervalSize;
                    var y1 = MARGIN + y * intervalSize;

                    var colour = Brushes.CornflowerBlue;
                    if (x == sourceIndexX && y == sourceIndexY)
                    {
                        colour = Brushes.OrangeRed;
                    }

                    if (x == targetIndexX && y == targetIndexY)
                    {
                        colour = Brushes.DarkOrange;
                    }

                    DrawRect(x1, y1, rectSize, rectSize, colour);
                }
            }
        }

        private void ComputeIndexes(out int sourceIndexX, out int sourceIndexY, out int targetIndexX, out int targetIndexY)
        {
            int xLast = XCopies - 1;
            int yLast = YCopies - 1;

            // Set X Axis Indexes
            if (_sourceLeft < _targetLeft)
            {
                sourceIndexX = 0;
                if (IsExtend)
                {
                    targetIndexX = 1;
                }
                else
                {
                    targetIndexX = xLast;
                }
            }
            else
            {
                sourceIndexX = xLast;
                if (IsExtend)
                {
                    targetIndexX = xLast - 1;
                }
                else
                {
                    targetIndexX = 0;
                }
            }

            // Set Y Axis Indexes
            if (_sourceTop < _targetTop)
            {
                sourceIndexY = 0;
                if (IsExtend)
                {
                    targetIndexY = 1;
                }
                else
                {
                    targetIndexY = yLast;
                }
            }
            else
            {
                sourceIndexY = yLast;
                if (IsExtend)
                {
                    targetIndexY = yLast - 1;
                }
                else
                {
                    targetIndexY = 0;
                }
            }
        }

        private void DrawLine(double x1, double y1, double x2, double y2)
        {
            var line = new Line
            {
                X1 = x1,
                Y1 = y1,
                X2 = x2,
                Y2 = y2,
                Stroke = Brushes.CornflowerBlue,
                StrokeThickness = 2
            };
            GridCanvas.Children.Add(line);
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
            GridCanvas.Children.Add(rect);
        }

        private void InitialiseDataSource()
        {
            dataSource = FindResource("DataSource") as DrawingsLabMultiCloneGridDataSource;
        }

        private void ButtomDialogOk_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }

        public int XCopies
        {
            get
            {
                if (dataSource.XCopies < MIN_COPIES)
                {
                    return MIN_COPIES;
                }
                return dataSource.XCopies;
            }
        }

        public int YCopies
        {
            get
            {
                if (dataSource.YCopies < MIN_COPIES)
                {
                    return MIN_COPIES;
                }
                return dataSource.YCopies;
            }
        }

        public bool IsExtend
        {
            get { return dataSource.IsExtend; }
        }
    }
}
