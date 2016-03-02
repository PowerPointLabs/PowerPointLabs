using System;
using System.Collections.Generic;
using System.Windows;
using PPExtraEventHelper;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using PowerPointLabs.ActionFramework.Common.Extension;

namespace PowerPointLabs.PositionsLab
{
    /// <summary>
    /// Interaction logic for PositionsPaneWPF.xaml
    /// </summary>
    public partial class PositionsPaneWPF : System.Windows.Controls.UserControl
    {
        private PositionsDistributeGridDialog positionsDistributeGridDialog;

        private static LMouseUpListener _leftMouseUpListener = null;
        private static LMouseDownListener _leftMouseDownListener = null;
        private static System.Windows.Threading.DispatcherTimer dispatcherTimer = new System.Windows.Threading.DispatcherTimer();

        //Variables for lock axis
        private const int LEFT = 0;
        private const int TOP = 1;
        private static bool isLockAxisMode = false;
        private static PowerPoint.ShapeRange shapesToBeMoved = null;
        private static System.Drawing.Point initialMousePos = new System.Drawing.Point();
        private float[,] initialPos;
        private static int timeCounter = 0;

        //Variables for rotation
        private const float REFPOINT_RADIUS = 10;
        private static bool isRotationMode = false;
        private static Shape refPoint = null;
        private static List<Shape> shapesToBeRotated = new List<Shape>();
        private static List<Shape> allShapesInSlide = new List<Shape>();
        private static System.Drawing.Point prevMousePos = new System.Drawing.Point();

        //Variables for settings
        private AlignSettingsDialog _alignSettingsDialog;
        private DistributeSettingsDialog _distributeSettingsDialog;

        public PositionsPaneWPF()
        {
            _alignSettingsDialog = new AlignSettingsDialog();
            _distributeSettingsDialog = new DistributeSettingsDialog();
            InitializeComponent();
            dispatcherTimer.Interval = TimeSpan.FromMilliseconds(10);
        }

        #region Align
        private void AlignLeftButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            List<Shape> selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
            PositionsLabMain.AlignLeft(selectedShapes);
        }

        private void AlignRightButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            List<Shape> selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            PositionsLabMain.AlignRight(selectedShapes, slideWidth);
        }

        private void AlignTopButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            List<Shape> selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
            PositionsLabMain.AlignTop(selectedShapes);
        }

        private void AlignBottomButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            List<Shape> selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            PositionsLabMain.AlignBottom(selectedShapes, slideHeight);
        }

        private void AlignMiddleButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            List<Shape> selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            PositionsLabMain.AlignMiddle(selectedShapes, slideHeight);
        }

        private void AlignCenterButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            List<Shape> selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            PositionsLabMain.AlignCenter(selectedShapes, slideWidth, slideHeight);
        }
        #endregion

        #region Adjoin
        private void AdjoinHorizontalButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            List<Shape> selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
            PositionsLabMain.AdjoinHorizontal(selectedShapes);
        }

        private void AdjoinVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            List<Shape> selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
            PositionsLabMain.AdjoinVertical(selectedShapes);
        }
        #endregion

        #region Distribute
        private void DistributeHorizontalButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            List<Shape> selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            PositionsLabMain.DistributeHorizontal(selectedShapes, slideWidth);
        }

        private void DistributeVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            List<Shape> selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            PositionsLabMain.DistributeVertical(selectedShapes, slideHeight);
        }

        private void DistributeCenterButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            List<Shape> selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            PositionsLabMain.DistributeCenter(selectedShapes, slideWidth, slideHeight);
        }

        private void DistributeShapesButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            List<Shape> selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
            PositionsLabMain.DistributeShapes(selectedShapes);
        }

        private void DistributeGridButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            List<Shape> selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
            int numShapesSelected = selectedShapes.Count;
            int rowLength = (int)Math.Ceiling(Math.Sqrt(numShapesSelected));
            int colLength = (int)Math.Ceiling((double)numShapesSelected / rowLength);

            if (positionsDistributeGridDialog == null || !positionsDistributeGridDialog.IsOpen)
            {
                positionsDistributeGridDialog = new PositionsDistributeGridDialog(selectedShapes, rowLength, colLength);
                positionsDistributeGridDialog.Show();
            }
            else
            {
                positionsDistributeGridDialog.Activate();
            }
        }
        #endregion

        #region Snap
        private void SnapHorizontalButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            List<Shape> selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
            PositionsLabMain.SnapHorizontal(selectedShapes);
        }

        private void SnapVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            List<Shape> selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
            PositionsLabMain.SnapVertical(selectedShapes);
        }

        private void SnapAwayButton_Click(object sender, RoutedEventArgs e)
        {
            bool noShapesSelected = this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes;

            if (noShapesSelected)
            {
                return;
            }

            PowerPoint.ShapeRange selectedShapes = this.GetCurrentSelection().ShapeRange;

            PositionsLabMain.SnapAway(ConvertShapeRangeToList(selectedShapes, 1));
        }
        #endregion

        #region Swap
        private void SwapPositionsButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            List<Shape> selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
            PositionsLabMain.Swap(selectedShapes);
        }
        #endregion

        #region Adjustment
        private void RotationButton_Click(object sender, RoutedEventArgs e)
        {
            bool noShapesSelected = this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes;

            if (noShapesSelected)
            {
                return;
            }

            PowerPoint.ShapeRange selectedShapes = this.GetCurrentSelection().ShapeRange;

            if (selectedShapes.Count <= 1)
            {
                return;
            }

            if (isLockAxisMode)
            {
                DisableLockAxisMode();
            }

            isRotationMode = true;

            var currentSlide = this.GetCurrentSlide();

            refPoint = selectedShapes[1];
            shapesToBeRotated = ConvertShapeRangeToList(selectedShapes, 2);
            allShapesInSlide = ConvertShapesToList(currentSlide.Shapes);

            dispatcherTimer.Tick += new EventHandler(RotationHandler);

            _leftMouseUpListener = new LMouseUpListener();
            _leftMouseUpListener.LButtonUpClicked +=
                new EventHandler<SysMouseEventInfo>(_leftMouseUpListener_Rotation);

            _leftMouseDownListener = new LMouseDownListener();
            _leftMouseDownListener.LButtonDownClicked +=
                new EventHandler<SysMouseEventInfo>(_leftMouseDownListener_Rotation);

        }

        private void RotationHandler(object sender, EventArgs e)
        {
            //Remove dragging control of user
            this.GetCurrentSelection().Unselect();
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

                    if (isLockAxisMode)
                    {
                        StartLockAxisMode();
                    }
                    return;
                }

                bool isShapeToBeRotated = shapesToBeRotated.Contains(selectedShape);
                bool isRefPoint = refPoint.Id == selectedShape.Id;

                if (!isShapeToBeRotated && !isRefPoint)
                {
                    DisableRotationMode();

                    if (isLockAxisMode)
                    {
                        StartLockAxisMode();
                    }
                    return;
                }

                if (isRefPoint)
                {
                    this.GetCurrentSelection().Unselect();
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
            DisableLockAxisMode();
            isLockAxisMode = false;
        }

        private void LockAxis_Checked(object sender, RoutedEventArgs e)
        {
            if (isRotationMode)
            {
                DisableRotationMode();
            }

            StartLockAxisMode();
            isLockAxisMode = true;
        }

        private void LockAxisHandler(object sender, EventArgs e)
        {
            //Allow mouseclick to register so that shape is selected
            //Solves bug where old selection replaces new selection due to _leftMouseUpListener_LockAxis
            if (timeCounter < 1)
            {
                timeCounter++;
                return;
            }

            bool noShapesSelected = this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes;

            if (shapesToBeMoved == null && noShapesSelected)
            {
                return;
            }

            if (shapesToBeMoved == null)
            {
                shapesToBeMoved = this.GetCurrentSelection().ShapeRange;
                initialPos = new float[shapesToBeMoved.Count, 2];
                for (int i = 0; i < shapesToBeMoved.Count; i++)
                {
                    Shape s = shapesToBeMoved[i + 1];
                    initialPos[i, LEFT] = s.Left;
                    initialPos[i, TOP] = s.Top;
                }
            }

            //Remove dragging control of user
            this.GetCurrentSelection().Unselect();

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
                var currentSlide = this.GetCurrentSlide();
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
                Logger.LogException(ex, "LockAxis");
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
            return this.GetCurrentWindow().PointsToScreenPixelsX(point);
        }

        private float PointsToScreenPixelsY(float point)
        {
            return this.GetCurrentWindow().PointsToScreenPixelsY(point);
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

        private List<Shape> ConvertShapeRangeToList (PowerPoint.ShapeRange range, int index)
        {
            List<Shape> shapes = new List<Shape>();

            for (int i = index; i <= range.Count; i++)
            {
                shapes.Add(range[i]);
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
        private void AlignSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            _alignSettingsDialog.ShowDialog();
        }

        private void DistributeSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            _distributeSettingsDialog.ShowDialog();            
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
            isRotationMode = false;

            if (refPoint != null)
            {
                refPoint = null;
            }

            shapesToBeRotated = new List<Shape>();
            allShapesInSlide = new List<Shape>();
            prevMousePos = new System.Drawing.Point();
        }

        private void StartLockAxisMode()
        {
            dispatcherTimer.Tick += new EventHandler(LockAxisHandler);

            _leftMouseUpListener = new LMouseUpListener();
            _leftMouseUpListener.LButtonUpClicked +=
                new EventHandler<SysMouseEventInfo>(_leftMouseUpListener_LockAxis);

            _leftMouseDownListener = new LMouseDownListener();
            _leftMouseDownListener.LButtonDownClicked +=
                new EventHandler<SysMouseEventInfo>(_leftMouseDownListener_LockAxis);
        }

        private static void DisableLockAxisMode()
        {
            ClearAllEventHandlers();
            shapesToBeMoved = null;
            initialMousePos = new System.Drawing.Point();
            timeCounter = 0;
        }
    }
}
