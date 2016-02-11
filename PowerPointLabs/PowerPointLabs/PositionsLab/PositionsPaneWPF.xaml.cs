using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MahApps.Metro.Controls;
using PPExtraEventHelper;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PositionsLab
{
    /// <summary>
    /// Interaction logic for PositionsPaneWPF.xaml
    /// </summary>
    public partial class PositionsPaneWPF : System.Windows.Controls.UserControl
    {

        private static LMouseUpListener _leftMouseUpListener = null;
        private static LMouseDownListener _leftMouseDownListener = null;

        private static System.Windows.Threading.DispatcherTimer dispatcherTimer = new System.Windows.Threading.DispatcherTimer();

        private const int LEFT = 0;
        private const int TOP = 1;
        private const int ROTATION = 2;

        //Variables for axis
        private const float REFPOINT_RADIUS = 10;
        private static Shape refPoint = null;
        private static List<Shape> shapesToBeRotated = new List<Shape>();
        private List<Shape> currentSelection = new List<Shape>();
        private float[,] initialShapes;
        private System.Drawing.Point prevMousePos = new System.Drawing.Point();

        public PositionsPaneWPF()
        {
            InitializeComponent();
            dispatcherTimer.Interval = TimeSpan.FromMilliseconds(10);
        }

        #region Align
        private void AlignLeftButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AlignLeft();
        }

        private void AlignRightButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AlignRight();
        }

        private void AlignTopButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AlignTop();
        }

        private void AlignBottomButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AlignBottom();
        }

        private void AlignMiddleButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AlignMiddle();
        }

        private void AlignCenterButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AlignCenter();
        }
        #endregion

        #region Adjoin
        private void AdjoinHorizontalButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AdjoinHorizontal();
        }

        private void AdjoinVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AdjoinVertical();
        }
        #endregion

        #region Distribute
        private void DistributeHorizontalButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.DistributeHorizontal();
        }

        private void DistributeVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.DistributeVertical();
        }

        private void DistributeCenterButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.DistributeCenter();
        }
        #endregion

        #region Snap
        private void SnapHorizontalButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.SnapHorizontal();
        }

        private void SnapVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.SnapVertical();
        }
        #endregion

        #region Swap
        private void SwapPositionsButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.Swap();
        }
        #endregion

        #region Adjustment
        private void RotationButton_Click(object sender, RoutedEventArgs e)
        {

            PowerPoint.ShapeRange selectedShapes = null;

            try
            {
                selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;
            }
            catch (Exception ex)
            {
                PowerPointLabsGlobals.LogException(ex, "RotationButtion_Click");
                return;
            }
            
            if (selectedShapes.Count > 0)
            {
                shapesToBeRotated = ConvertShapeRangeToList(selectedShapes);
                System.Drawing.PointF refCoordinates = CalculateCenterPoint(shapesToBeRotated);
                refPoint = AddReferencePoint(Globals.ThisAddIn.Application.ActiveWindow.View.Slide.Shapes, refCoordinates.X - REFPOINT_RADIUS/2, refCoordinates.Y - REFPOINT_RADIUS/2);

                dispatcherTimer.Tick += new EventHandler(RotationHandler);

                _leftMouseUpListener = new LMouseUpListener();
                _leftMouseUpListener.LButtonUpClicked +=
                    new EventHandler<SysMouseEventInfo>(_leftMouseUpListener_Rotation);

                _leftMouseDownListener = new LMouseDownListener();
                _leftMouseDownListener.LButtonDownClicked +=
                    new EventHandler<SysMouseEventInfo>(_leftMouseDownListener_Rotation);

            }
        }

        private void RotationHandler(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            System.Drawing.Point p = System.Windows.Forms.Control.MousePosition;

            float prevAngle = (float)AngleBetweenTwoPoints(ConvertSlidePointToScreenPoint(Graphics.GetCenterPoint(refPoint)), prevMousePos);
            float angle = (float)AngleBetweenTwoPoints(ConvertSlidePointToScreenPoint(Graphics.GetCenterPoint(refPoint)), p) - prevAngle;
            System.Drawing.PointF origin = Graphics.GetCenterPoint(refPoint);

            for (int i = 0; i < currentSelection.Count; i++)
            {
                Shape currentShape = currentSelection[i];
                System.Drawing.PointF unrotatedCenter = Graphics.GetCenterPoint(currentShape);
                System.Drawing.PointF rotatedCenter = Graphics.RotatePoint(unrotatedCenter, origin, angle);

                currentShape.Left += (rotatedCenter.X - unrotatedCenter.X);
                currentShape.Top += (rotatedCenter.Y - unrotatedCenter.Y);

                currentShape.Rotation = AddAngles(currentShape.Rotation, angle);
            }

            prevMousePos = p;
        }

        void _leftMouseUpListener_Rotation(object sender, SysMouseEventInfo e)
        {
            dispatcherTimer.Stop();
        }

        void _leftMouseDownListener_Rotation(object sender, SysMouseEventInfo e)
        {
            try
            {
                System.Drawing.Point p = System.Windows.Forms.Control.MousePosition;

                Shape selectedShape = ShapeBelowMousePos(shapesToBeRotated, p);

                if (selectedShape == null && !IsPointWithinShape(refPoint, p))
                {
                    ResetRotationMode();
                    return;
                }

                if (IsPointWithinShape(refPoint, p))
                {
                    return;
                }

                selectedShape.Select();
                currentSelection = ConvertShapeRangeToList(Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange);
                prevMousePos = System.Windows.Forms.Control.MousePosition;
                initialShapes = new float[currentSelection.Count, 3];

                for (int i = 0; i < currentSelection.Count; i++)
                {
                    initialShapes[i, LEFT] = currentSelection[i].Left;
                    initialShapes[i, TOP] = currentSelection[i].Top;
                    initialShapes[i, ROTATION] = currentSelection[i].Rotation;
                }

                dispatcherTimer.Start();
            }
            catch (Exception ex)
            {
                PowerPointLabsGlobals.LogException(ex, "LockAxis");
            }
        }
        #endregion

        #region Helper
        private Shape AddReferencePoint(PowerPoint.Shapes shapes, float left, float top)
        {
            return shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, left, top, REFPOINT_RADIUS, REFPOINT_RADIUS);
        }

        private System.Drawing.PointF CalculateCenterPoint(List<Shape> shapes)
        {
            System.Drawing.PointF centerPoint = new System.Drawing.PointF();

            if (shapes.Count == 1)
            {
                centerPoint = Graphics.GetCenterPoint(shapes[0]);
            }

            foreach (Shape s in shapes)
            {
                System.Drawing.PointF currentShapeCenterPoint = Graphics.GetCenterPoint(s);
                centerPoint.X = (centerPoint.X + currentShapeCenterPoint.X) / 2;
                centerPoint.Y = (centerPoint.Y + currentShapeCenterPoint.Y) / 2;
            }

            return centerPoint;
        }

        private float PointsToScreenPixelsX(float point)
        {
            return Globals.ThisAddIn.Application.ActiveWindow.PointsToScreenPixelsX(point);
        }

        private float PointsToScreenPixelsY(float point)
        {
            return Globals.ThisAddIn.Application.ActiveWindow.PointsToScreenPixelsY(point);
        }

        private bool IsPointWithinShape(Shape shape, System.Drawing.Point p)
        {
            float x1 = PointsToScreenPixelsX(shape.Left);
            float y1 = PointsToScreenPixelsY(shape.Top);
            float x2 = PointsToScreenPixelsX(shape.Left + shape.Width);
            float y2 = PointsToScreenPixelsY(shape.Top + shape.Height);

            // Expand the bounding box with a standard padding
            // NOTE: PowerPoint transforms the mouse cursor when entering shapes before it actually
            // enters the shape. To account for that, add this extra 'padding'
            // Testing reveals that the current value (in PowerPoint 2013) is 6px
            // http://stackoverflow.com/questions/22815084/catch-mouse-events-in-powerpoint-designer-through-vsto
            x1 -= 6;
            x2 += 6;
            y1 -= 6;
            y2 += 6;

            return (x1 <= p.X && p.X <= x2) && (y1 <= p.Y && p.Y <= y2);
        }

        private Shape ShapeBelowMousePos(List<Shape> shapes, System.Drawing.Point p)
        {
            foreach (Shape s in shapes)
            {
                if (IsPointWithinShape(s, p))
                {
                    return s;
                }
            }

            return null;
        }

        private List<Shape> ConvertShapeRangeToList (PowerPoint.ShapeRange range)
        {
            List<Shape> shapes = new List<Shape>();

            foreach (Shape s in range)
            {
                shapes.Add(s);
            }

            return shapes;
        }

        private double AngleBetweenTwoPoints(System.Drawing.PointF refPoint, System.Drawing.PointF pt)
        {
            double angle = Math.Atan((pt.Y - refPoint.Y) / (pt.X - refPoint.X)) * 180 / Math.PI;

            if (pt.X - refPoint.X > 0)
            {
                angle = 90 + angle;
            }
            else
            {
                angle = 270 + angle;
            }

            return angle;
        }

        private System.Drawing.PointF ConvertSlidePointToScreenPoint(System.Drawing.PointF pt)
        {
            pt.X = PointsToScreenPixelsX(pt.X);
            pt.Y = PointsToScreenPixelsX(pt.Y);

            return pt;
        }

        private float AddAngles(float a, float b)
        {
            return (a + b) % 360;
        }

        #endregion

        public static void ClearAllEventHandlers()
        {
            if (_leftMouseUpListener != null)
            {
                _leftMouseUpListener.Close();
            }

            if (_leftMouseDownListener != null)
            {
                _leftMouseDownListener.Close();
            }

            dispatcherTimer = new System.Windows.Threading.DispatcherTimer();
        }

        private static void ResetRotationMode()
        {
            ClearAllEventHandlers();

            if (refPoint != null)
            {
                refPoint.Delete();
                refPoint = null;
            }

            shapesToBeRotated = new List<Shape>();
        }
    }
}
