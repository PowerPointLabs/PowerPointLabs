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
    /// Interaction logic for AlignmentDialogHorizontal.xaml
    /// </summary>
    public partial class AlignmentDialogHorizontal : Window
    {
        private DrawingsLabAlignmentDataSource dataSource;
        
        public AlignmentDialogHorizontal()
        {
            InitializeComponent();

            InitialiseDataSource();

            dataSource.targetPropertyChangeEvent += DrawAlignmentCanvas;
            dataSource.sourcePropertyChangeEvent += DrawAlignmentCanvas;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            DrawAlignmentCanvas();
        }

        private double CanvasAbsoluteY(double f)
        {
            return f * AlignmentCanvas.ActualHeight;
        }

        private void DrawAlignmentCanvas()
        {
            AlignmentCanvas.Children.Clear();
            double middle = AlignmentCanvas.ActualWidth / 2;
            double gapWidth = 10f;

            double targetSquareWidth = CanvasAbsoluteY(1f / 3f);
            double sourceSquareWidth = CanvasAbsoluteY(1f / 4f);
            double anchorY = (100-TargetAnchor) / 300f + 1 / 3f;
            double topY = anchorY - (100-SourceAnchor) / 400f;

            DrawRect(middle + gapWidth, CanvasAbsoluteY(1f / 3), targetSquareWidth, targetSquareWidth, Brushes.OrangeRed);
            DrawRect(middle - gapWidth - sourceSquareWidth, CanvasAbsoluteY(topY), sourceSquareWidth, sourceSquareWidth, Brushes.DarkOrange);

            var line = new Line
            {
                X1 = middle - gapWidth - sourceSquareWidth - 10f,
                Y1 = CanvasAbsoluteY(anchorY),
                X2 = middle + gapWidth + targetSquareWidth + 10f,
                Y2 = CanvasAbsoluteY(anchorY),
                Stroke = Brushes.CornflowerBlue,
                StrokeThickness = 2
            };
            AlignmentCanvas.Children.Add(line);
        }

        private void DrawRect(double x, double y, double width, double height, Brush colour)
        {
            var rect = new Rectangle
            {
                Width = width,
                Height = height,
                Stroke = colour,
                StrokeThickness = 3
            };
            Canvas.SetLeft(rect, x);
            Canvas.SetTop(rect, y);
            AlignmentCanvas.Children.Add(rect);
        }


        private void InitialiseDataSource()
        {
            dataSource = FindResource("DataSource") as DrawingsLabAlignmentDataSource;
        }

        private void ButtomDialogOk_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }

        public double SourceAnchor
        {
            get { return dataSource.SourceAnchor; }
        }

        public double TargetAnchor
        {
            get { return dataSource.TargetAnchor; }
        }
    }

}
