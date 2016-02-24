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
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using PowerPointLabs.Models;

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

        //Variables for lock axis
        private const int LEFT = 0;
        private const int TOP = 1;
        private PowerPoint.ShapeRange shapesToBeMoved = null;
        private static System.Drawing.Point initialMousePos = new System.Drawing.Point();
        private float[,] initialPos;
        private int timeCounter = 0;

        //Variables for rotation
        private const float REFPOINT_RADIUS = 10;
        private static Shape refPoint = null;
        private static List<Shape> shapesToBeRotated = new List<Shape>();
        private static List<Shape> allShapesInSlide = new List<Shape>();
        private static System.Drawing.Point prevMousePos = new System.Drawing.Point();

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

        private void DistributeShapesButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.DistributeShapes();
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

        private void SnapAwayButton_Click(object sender, RoutedEventArgs e)
        {
            bool noShapesSelected = Globals.ThisAddIn.Application.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes;

            if (noShapesSelected)
            {
                return;
            }

            PowerPoint.ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

            PositionsLabMain.SnapAway(ConvertShapeRangeToList(selectedShapes));
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
            bool noShapesSelected = Globals.ThisAddIn.Application.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes;

            if (noShapesSelected)
            {
                return;
            }

            PowerPoint.ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

            if (selectedShapes.Count > 0)
            {
                var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide as PowerPointSlide;

                shapesToBeRotated = ConvertShapeRangeToList(selectedShapes);
                System.Drawing.PointF refCoordinates = CalculateCenterPoint(shapesToBeRotated);
                refPoint = AddReferencePoint(currentSlide.Shapes, refCoordinates.X - REFPOINT_RADIUS/2, refCoordinates.Y - REFPOINT_RADIUS/2);
                allShapesInSlide = ConvertShapesToList(currentSlide.Shapes);

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
            //Remove dragging control of user
            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            System.Drawing.Point p = System.Windows.Forms.Control.MousePosition;

            float prevAngle = (float)PositionsLabMain.AngleBetweenTwoPoints(ConvertSlidePointToScreenPoint(Graphics.GetCenterPoint(refPoint)), prevMousePos);
            float angle = (float)PositionsLabMain.AngleBetweenTwoPoints(ConvertSlidePointToScreenPoint(Graphics.GetCenterPoint(refPoint)), p) - prevAngle;
            System.Drawing.PointF origin = Graphics.GetCenterPoint(refPoint);

            for (int i = 0; i < shapesToBeRotated.Count; i++)
            {
                Shape currentShape = shapesToBeRotated[i];
                System.Drawing.PointF unrotatedCenter = Graphics.GetCenterPoint(currentShape);
                System.Drawing.PointF rotatedCenter = Graphics.RotatePoint(unrotatedCenter, origin, angle);

                currentShape.Left += (rotatedCenter.X - unrotatedCenter.X);
                currentShape.Top += (rotatedCenter.Y - unrotatedCenter.Y);

                currentShape.Rotation = PositionsLabMain.AddAngles(currentShape.Rotation, angle);
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
                Shape selectedShape = GetShapeDirectlyBelowMousePos(allShapesInSlide, p);

                if (selectedShape == null)
                {
                    DisableRotationMode();
                    return;
                }

                bool isShapeToBeRotated = shapesToBeRotated.Contains(selectedShape);
                bool isRefPoint = refPoint.Id == selectedShape.Id;

                if (!isShapeToBeRotated && !isRefPoint)
                {
                    DisableRotationMode();
                    return;
                }

                if (isRefPoint)
                {
                    return;
                }

                prevMousePos = p;
                dispatcherTimer.Start();
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "Rotation");
            }
        }

        private void LockAxis_UnChecked(object sender, RoutedEventArgs e)
        {
            ClearAllEventHandlers();
            shapesToBeMoved = null;
            initialMousePos = new System.Drawing.Point();
            timeCounter = 0;
        }

        private void LockAxis_Checked(object sender, RoutedEventArgs e)
        {
            dispatcherTimer.Tick += new EventHandler(LockAxisHandler);

            _leftMouseUpListener = new LMouseUpListener();
            _leftMouseUpListener.LButtonUpClicked +=
                new EventHandler<SysMouseEventInfo>(_leftMouseUpListener_LockAxis);

            _leftMouseDownListener = new LMouseDownListener();
            _leftMouseDownListener.LButtonDownClicked +=
                new EventHandler<SysMouseEventInfo>(_leftMouseDownListener_LockAxis);
        }

        private void LockAxisHandler(object sender, EventArgs e)
        {
            if (timeCounter < 1)
            {
                timeCounter++;
                return;
            }

            bool noShapesSelected = Globals.ThisAddIn.Application.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes;

            if (shapesToBeMoved == null && noShapesSelected)
            {
                return;
            }

            if (shapesToBeMoved == null)
            {
                shapesToBeMoved = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                initialPos = new float[shapesToBeMoved.Count, 2];
                for (int i = 0; i < shapesToBeMoved.Count; i++)
                {
                    Shape s = shapesToBeMoved[i + 1];
                    initialPos[i, LEFT] = s.Left;
                    initialPos[i, TOP] = s.Top;
                }
            }

            //Remove dragging control of user
            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();

            System.Drawing.Point currentMousePos = System.Windows.Forms.Control.MousePosition;

            float diffX = currentMousePos.X - initialMousePos.X;
            float diffY = currentMousePos.Y - initialMousePos.Y;

            for (int i = 0; i < shapesToBeMoved.Count; i++)
            {
                Shape s = shapesToBeMoved[i + 1];
                if (Math.Abs(diffX) > Math.Abs(diffY))
                {
                    s.Left = initialPos[i, LEFT] + diffX;
                    s.Top = initialPos[i, TOP];
                }
                else
                {
                    s.Left = initialPos[i, LEFT];
                    s.Top = initialPos[i, TOP] + diffY;
                }
            }
        }

        void _leftMouseUpListener_LockAxis(object sender, SysMouseEventInfo e)
        {
            dispatcherTimer.Stop();
            timeCounter = 0;
            if (shapesToBeMoved != null)
            {
                shapesToBeMoved.Select();
                shapesToBeMoved = null;
            }
        }

        void _leftMouseDownListener_LockAxis(object sender, SysMouseEventInfo e)
        {
            try
            {
                System.Drawing.Point p = System.Windows.Forms.Control.MousePosition;
                var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide as PowerPointSlide;
                allShapesInSlide = ConvertShapesToList(currentSlide.Shapes);
                Shape selectedShape = GetShapeDirectlyBelowMousePos(allShapesInSlide, p);

                if (selectedShape == null || PPKeyboard.IsCtrlPressed() || PPKeyboard.IsShiftPressed())
                {
                    return;
                }

                initialMousePos = p;
                dispatcherTimer.Start();
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "Rotation");
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
            if (shapes.Count < 1)
            {
                return new System.Drawing.PointF();
            }

            System.Drawing.PointF centerPoint = Graphics.GetCenterPoint(shapes[0]);

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
            float epsilon = 0.00001f;

            System.Drawing.PointF centerPoint = ConvertSlidePointToScreenPoint(Graphics.GetCenterPoint(shape));
            System.Drawing.PointF rotatedMousePos = Graphics.RotatePoint(p, centerPoint, -shape.Rotation);

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

            return (x1 - epsilon <= rotatedMousePos.X && rotatedMousePos.X  <= x2 + epsilon) && (y1 - epsilon <= rotatedMousePos.Y && rotatedMousePos.Y <= y2 + epsilon);
        }

        private Shape GetShapeDirectlyBelowMousePos(List<Shape> shapes, System.Drawing.Point p)
        {
            Shape aShape = null;

            foreach (Shape s in shapes)
            {
                if (IsPointWithinShape(s, p))
                {
                    if (aShape == null || aShape.ZOrderPosition < s.ZOrderPosition)
                    {
                        aShape = s;
                    }
                }
            }

            return aShape;
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

        private List<Shape> ConvertShapesToList(PowerPoint.Shapes shapes)
        {
            List<Shape> listOfShapes = new List<Shape>();

            foreach (Shape s in shapes)
            {
                listOfShapes.Add(s);
            }

            return listOfShapes;
        }

        private System.Drawing.PointF ConvertSlidePointToScreenPoint(System.Drawing.PointF pt)
        {
            pt.X = PointsToScreenPixelsX(pt.X);
            pt.Y = PointsToScreenPixelsY(pt.Y);

            return pt;
        }

        private void SelectShapes(List<Shape> shapes)
        {
            foreach (Shape s in shapes)
            {
                s.Select(Office.MsoTriState.msoFalse);
            }
        }

        #endregion

        #region Settings
        // Note: if changing default behavior to using slide as reference, need to ensure that
        // checkbox for using shape is defined first in PositionsPaneWPF.xaml

        // TODO: Surround with try catch in case the order of checkboxes are wrong
        private void UseShapeAsReference(object sender, RoutedEventArgs e)
        {
            if (!slideAsReference.IsChecked.HasValue || !shapeAsReference.IsChecked.HasValue)
            {
                //Error
                return;
            }
            slideAsReference.IsChecked = false;
            shapeAsReference.IsChecked = true;
            PositionsLabMain.ReferToShape();
        }

        private void UseSlideAsReference(object sender, RoutedEventArgs e)
        {
            if (!slideAsReference.IsChecked.HasValue || !shapeAsReference.IsChecked.HasValue)
            {
                //Error
                return;
            }
            shapeAsReference.IsChecked = false;
            slideAsReference.IsChecked = true;
            PositionsLabMain.ReferToSlide();
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

            dispatcherTimer.Stop();
            dispatcherTimer = new System.Windows.Threading.DispatcherTimer();
        }

        private static void DisableRotationMode()
        {
            ClearAllEventHandlers();

            if (refPoint != null)
            {
                refPoint.Delete();
                refPoint = null;
            }

            shapesToBeRotated = new List<Shape>();
            allShapesInSlide = new List<Shape>();
            prevMousePos = new System.Drawing.Point();
        }

    }
}
