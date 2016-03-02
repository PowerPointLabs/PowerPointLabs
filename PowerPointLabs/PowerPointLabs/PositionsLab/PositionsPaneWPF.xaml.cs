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
    public partial class PositionsPaneWpf
    {
        private PositionsDistributeGridDialog _positionsDistributeGridDialog;

        private static LMouseUpListener _leftMouseUpListener;
        private static LMouseDownListener _leftMouseDownListener;
        private static System.Windows.Threading.DispatcherTimer _dispatcherTimer = new System.Windows.Threading.DispatcherTimer();

        //Variables for lock axis
        private const int Left = 0;
        private const int Top = 1;
        private static bool _isLockAxisMode;
        private static PowerPoint.ShapeRange _shapesToBeMoved;
        private static System.Drawing.Point _initialMousePos;
        private float[,] _initialPos;
        private static int _timeCounter;

        //Variables for rotation
        private const float RefpointRadius = 10;
        private static bool _isRotationMode;
        private static Shape _refPoint;
        private static List<Shape> _shapesToBeRotated = new List<Shape>();
        private static List<Shape> _allShapesInSlide = new List<Shape>();
        private static System.Drawing.Point _prevMousePos;

        //Variables for settings
        private AlignSettingsDialog _alignSettingsDialog;
        private DistributeSettingsDialog _distributeSettingsDialog;

        public PositionsPaneWpf()
        {
            _alignSettingsDialog = new AlignSettingsDialog();
            _distributeSettingsDialog = new DistributeSettingsDialog();
            InitializeComponent();
            _dispatcherTimer.Interval = TimeSpan.FromMilliseconds(10);
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

            if (_positionsDistributeGridDialog == null || !_positionsDistributeGridDialog.IsOpen)
            {
                _positionsDistributeGridDialog = new PositionsDistributeGridDialog(selectedShapes, rowLength, colLength);
                _positionsDistributeGridDialog.Show();
            }
            else
            {
                _positionsDistributeGridDialog.Activate();
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

            if (_isLockAxisMode)
            {
                DisableLockAxisMode();
            }

            _isRotationMode = true;

            var currentSlide = this.GetCurrentSlide();

            _refPoint = selectedShapes[1];
            _shapesToBeRotated = ConvertShapeRangeToList(selectedShapes, 2);
            _allShapesInSlide = ConvertShapesToList(currentSlide.Shapes);

            _dispatcherTimer.Tick += RotationHandler;

            _leftMouseUpListener = new LMouseUpListener();
            _leftMouseUpListener.LButtonUpClicked += _leftMouseUpListener_Rotation;

            _leftMouseDownListener = new LMouseDownListener();
            _leftMouseDownListener.LButtonDownClicked += _leftMouseDownListener_Rotation;

        }

        private void RotationHandler(object sender, EventArgs e)
        {
            //Remove dragging control of user
            this.GetCurrentSelection().Unselect();
            System.Drawing.Point p = System.Windows.Forms.Control.MousePosition;

            float prevAngle = (float)PositionsLabMain.AngleBetweenTwoPoints(ConvertSlidePointToScreenPoint(Graphics.GetCenterPoint(_refPoint)), _prevMousePos);
            float angle = (float)PositionsLabMain.AngleBetweenTwoPoints(ConvertSlidePointToScreenPoint(Graphics.GetCenterPoint(_refPoint)), p) - prevAngle;
            System.Drawing.PointF origin = Graphics.GetCenterPoint(_refPoint);

            foreach (Shape currentShape in _shapesToBeRotated)
            {
                System.Drawing.PointF unrotatedCenter = Graphics.GetCenterPoint(currentShape);
                System.Drawing.PointF rotatedCenter = Graphics.RotatePoint(unrotatedCenter, origin, angle);

                currentShape.Left += (rotatedCenter.X - unrotatedCenter.X);
                currentShape.Top += (rotatedCenter.Y - unrotatedCenter.Y);

                currentShape.Rotation = PositionsLabMain.AddAngles(currentShape.Rotation, angle);
            }

            _prevMousePos = p;
        }

        void _leftMouseUpListener_Rotation(object sender, SysMouseEventInfo e)
        {
            _dispatcherTimer.Stop();
        }

        void _leftMouseDownListener_Rotation(object sender, SysMouseEventInfo e)
        {
            try
            {
                System.Drawing.Point p = System.Windows.Forms.Control.MousePosition;
                Shape selectedShape = GetShapeDirectlyBelowMousePos(_allShapesInSlide, p);

                if (selectedShape == null)
                {
                    DisableRotationMode();

                    if (_isLockAxisMode)
                    {
                        StartLockAxisMode();
                    }
                    return;
                }

                bool isShapeToBeRotated = _shapesToBeRotated.Contains(selectedShape);
                bool isRefPoint = _refPoint.Id == selectedShape.Id;

                if (!isShapeToBeRotated && !isRefPoint)
                {
                    DisableRotationMode();

                    if (_isLockAxisMode)
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

                _prevMousePos = p;
                _dispatcherTimer.Start();
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "Rotation");
            }
        }

        private void LockAxis_UnChecked(object sender, RoutedEventArgs e)
        {
            DisableLockAxisMode();
            _isLockAxisMode = false;
        }

        private void LockAxis_Checked(object sender, RoutedEventArgs e)
        {
            if (_isRotationMode)
            {
                DisableRotationMode();
            }

            StartLockAxisMode();
            _isLockAxisMode = true;
        }

        private void LockAxisHandler(object sender, EventArgs e)
        {
            //Allow mouseclick to register so that shape is selected
            //Solves bug where old selection replaces new selection due to _leftMouseUpListener_LockAxis
            if (_timeCounter < 1)
            {
                _timeCounter++;
                return;
            }

            bool noShapesSelected = this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes;

            if (_shapesToBeMoved == null && noShapesSelected)
            {
                return;
            }

            if (_shapesToBeMoved == null)
            {
                _shapesToBeMoved = this.GetCurrentSelection().ShapeRange;
                _initialPos = new float[_shapesToBeMoved.Count, 2];
                for (int i = 0; i < _shapesToBeMoved.Count; i++)
                {
                    Shape s = _shapesToBeMoved[i + 1];
                    _initialPos[i, Left] = s.Left;
                    _initialPos[i, Top] = s.Top;
                }
            }

            //Remove dragging control of user
            this.GetCurrentSelection().Unselect();

            System.Drawing.Point currentMousePos = System.Windows.Forms.Control.MousePosition;

            float diffX = currentMousePos.X - _initialMousePos.X;
            float diffY = currentMousePos.Y - _initialMousePos.Y;

            for (int i = 0; i < _shapesToBeMoved.Count; i++)
            {
                Shape s = _shapesToBeMoved[i + 1];
                if (Math.Abs(diffX) > Math.Abs(diffY))
                {
                    s.Left = _initialPos[i, Left] + diffX;
                    s.Top = _initialPos[i, Top];
                }
                else
                {
                    s.Left = _initialPos[i, Left];
                    s.Top = _initialPos[i, Top] + diffY;
                }
            }
        }

        void _leftMouseUpListener_LockAxis(object sender, SysMouseEventInfo e)
        {
            _dispatcherTimer.Stop();
            _timeCounter = 0;
            if (_shapesToBeMoved != null)
            {
                _shapesToBeMoved.Select();
                _shapesToBeMoved = null;
            }
        }

        void _leftMouseDownListener_LockAxis(object sender, SysMouseEventInfo e)
        {
            try
            {
                System.Drawing.Point p = System.Windows.Forms.Control.MousePosition;
                var currentSlide = this.GetCurrentSlide();
                _allShapesInSlide = ConvertShapesToList(currentSlide.Shapes);
                Shape selectedShape = GetShapeDirectlyBelowMousePos(_allShapesInSlide, p);

                if (selectedShape == null || PPKeyboard.IsCtrlPressed() || PPKeyboard.IsShiftPressed())
                {
                    return;
                }

                _initialMousePos = p;
                _dispatcherTimer.Start();
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
            return shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left, top, RefpointRadius, RefpointRadius);
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

            _dispatcherTimer.Stop();
            _dispatcherTimer = new System.Windows.Threading.DispatcherTimer();
        }

        private static void DisableRotationMode()
        {
            ClearAllEventHandlers();
            _isRotationMode = false;

            if (_refPoint != null)
            {
                _refPoint = null;
            }

            _shapesToBeRotated = new List<Shape>();
            _allShapesInSlide = new List<Shape>();
            _prevMousePos = new System.Drawing.Point();
        }

        private void StartLockAxisMode()
        {
            _dispatcherTimer.Tick += LockAxisHandler;

            _leftMouseUpListener = new LMouseUpListener();
            _leftMouseUpListener.LButtonUpClicked += _leftMouseUpListener_LockAxis;

            _leftMouseDownListener = new LMouseDownListener();
            _leftMouseDownListener.LButtonDownClicked += _leftMouseDownListener_LockAxis;
        }

        private static void DisableLockAxisMode()
        {
            ClearAllEventHandlers();
            _shapesToBeMoved = null;
            _initialMousePos = new System.Drawing.Point();
            _timeCounter = 0;
        }
    }
}
