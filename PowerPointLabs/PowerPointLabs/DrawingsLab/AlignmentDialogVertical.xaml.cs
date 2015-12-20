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
    public partial class AlignmentDialogVertical : Window
    {
        private DrawingsLabAlignmentDataSource dataSource;

        public AlignmentDialogVertical()
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

        private float CanvasAbsoluteX(float f)
        {
            return (float)(f * AlignmentCanvas.ActualWidth);
        }

        private void DrawAlignmentCanvas()
        {
            AlignmentCanvas.Children.Clear();
            float middle = (float)(AlignmentCanvas.ActualHeight/2f);
            float gapHeight = 10f;

            float targetSquareWidth = CanvasAbsoluteX(1f / 3f);
            float sourceSquareWidth = CanvasAbsoluteX(1f / 4f);
            float anchorX = dataSource.TargetAnchor/300f + 1/3f;
            float leftX = anchorX - dataSource.SourceAnchor/400f;

            DrawRect(CanvasAbsoluteX(1f / 3), middle + gapHeight, targetSquareWidth, targetSquareWidth, Brushes.OrangeRed);
            DrawRect(CanvasAbsoluteX(leftX), middle - gapHeight - sourceSquareWidth, sourceSquareWidth, sourceSquareWidth, Brushes.DarkOrange);

            var line = new Line
            {
                X1 = CanvasAbsoluteX(anchorX),
                Y1 = middle - gapHeight - sourceSquareWidth - 10f,
                X2 = CanvasAbsoluteX(anchorX),
                Y2 = middle + gapHeight + targetSquareWidth + 10f,
                Stroke = Brushes.CornflowerBlue,
                StrokeThickness = 2
            };
            AlignmentCanvas.Children.Add(line);
        }

        private void DrawRect(float x, float y, float width, float height, Brush colour)
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

        public float SourceAnchor
        {
            get { return dataSource.SourceAnchor; }
        }

        public float TargetAnchor
        {
            get { return dataSource.TargetAnchor; }
        }
    }
}
