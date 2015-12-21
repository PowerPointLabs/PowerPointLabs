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
    /// Interaction logic for AlignmentDialogBoth.xaml
    /// </summary>
    public partial class AlignmentDialogBoth : Window
    {
        private DrawingsLabAlignmentDataSource dataSourceHorizontal;
        private DrawingsLabAlignmentDataSource dataSourceVertical;
        
        public AlignmentDialogBoth()
        {
            InitializeComponent();

            InitialiseDataSource();

            dataSourceHorizontal.targetPropertyChangeEvent += DrawAlignmentCanvas;
            dataSourceHorizontal.sourcePropertyChangeEvent += DrawAlignmentCanvas;
            dataSourceVertical.targetPropertyChangeEvent += DrawAlignmentCanvas;
            dataSourceVertical.sourcePropertyChangeEvent += DrawAlignmentCanvas;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            DrawAlignmentCanvas();
        }

        private float CanvasAbsoluteX(float f)
        {
            // Assumption: AlignmentCanvas.ActualWidth > AlignmentCanvasActualHeight

            var left = (AlignmentCanvas.ActualWidth - AlignmentCanvas.ActualHeight)/2;
            return (float)(left + f * AlignmentCanvas.ActualHeight);
        }

        private float CanvasAbsoluteY(float f)
        {
            return (float)(f * AlignmentCanvas.ActualHeight);
        }

        private void DrawAlignmentCanvas()
        {
            AlignmentCanvas.Children.Clear();
            float middleX = CanvasAbsoluteX(0.5f);
            float middleY = CanvasAbsoluteY(0.5f);

            float targetSquareWidth = CanvasAbsoluteY(1f / 3f);
            float sourceSquareWidth = CanvasAbsoluteY(1f / 4f);
            float lineHalfLength = targetSquareWidth/2 + sourceSquareWidth + 10f;;

            float anchorX = TargetAnchorVertical / 300f + 1 / 3f;
            float leftX = anchorX - SourceAnchorVertical / 400f;

            float anchorY = (100 - TargetAnchorHorizontal) / 300f + 1 / 3f;
            float topY = anchorY - (100 - SourceAnchorHorizontal) / 400f;

            DrawRect(middleX - targetSquareWidth/2, middleY - targetSquareWidth/2, targetSquareWidth, targetSquareWidth, Brushes.OrangeRed);
            DrawRect(CanvasAbsoluteX(leftX), CanvasAbsoluteY(topY), sourceSquareWidth, sourceSquareWidth, Brushes.DarkOrange);

            DrawLine(middleX - lineHalfLength, CanvasAbsoluteY(anchorY),
                     middleX + lineHalfLength, CanvasAbsoluteY(anchorY));
            DrawLine(CanvasAbsoluteX(anchorX), middleY - lineHalfLength,
                     CanvasAbsoluteX(anchorX), middleY + lineHalfLength);
        }

        private void DrawLine(float x1, float y1, float x2, float y2)
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
            dataSourceHorizontal = FindResource("DataSourceHorizontal") as DrawingsLabAlignmentDataSource;
            dataSourceVertical = FindResource("DataSourceVertical") as DrawingsLabAlignmentDataSource;
        }

        private void ButtomDialogOk_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }

        public float SourceAnchorHorizontal
        {
            get { return dataSourceHorizontal.SourceAnchor; }
        }

        public float TargetAnchorHorizontal
        {
            get { return dataSourceHorizontal.TargetAnchor; }
        }

        public float SourceAnchorVertical
        {
            get { return dataSourceVertical.SourceAnchor; }
        }

        public float TargetAnchorVertical
        {
            get { return dataSourceVertical.TargetAnchor; }
        }
    }

}
