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
using Graphics = PowerPointLabs.Utils.Graphics;
using System.Diagnostics;

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

        //Error Messages
        private const string ErrorMessageNoSelection = TextCollection.PositionsLabText.ErrorNoSelection;
        private const string ErrorMessageFewerThanTwoSelection = TextCollection.PositionsLabText.ErrorFewerThanTwoSelection;
        private const string ErrorMessageFewerThanThreeSelection =
            TextCollection.PositionsLabText.ErrorFewerThanThreeSelection;
        private const string ErrorMessageUndefined = TextCollection.PositionsLabText.ErrorUndefined;

        //Variable for preview
        bool _previewIsExecuted = false;

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
        private ReorderSettingsDialog _reorderSettingsDialog;

        public PositionsPaneWpf()
        {
            PositionsLabMain.InitPositionsLab();
            InitializeComponent();
            _dispatcherTimer.Interval = TimeSpan.FromMilliseconds(10);
        }

        #region Click Behaviour
        #region Align
        private void AlignLeftButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                PositionsLabMain.AlignLeft(selectedShapes);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }

        private void AlignRightButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                var slideWidth = this.GetCurrentPresentation().SlideWidth;
                PositionsLabMain.AlignRight(selectedShapes, slideWidth);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }

        private void AlignTopButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                PositionsLabMain.AlignTop(selectedShapes);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }

        private void AlignBottomButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                var slideHeight = this.GetCurrentPresentation().SlideHeight;
                PositionsLabMain.AlignBottom(selectedShapes, slideHeight);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }

        private void AlignMiddleButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                var slideHeight = this.GetCurrentPresentation().SlideHeight;
                PositionsLabMain.AlignMiddle(selectedShapes, slideHeight);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }

        private void AlignCenterButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                var slideHeight = this.GetCurrentPresentation().SlideHeight;
                var slideWidth = this.GetCurrentPresentation().SlideWidth;
                PositionsLabMain.AlignCenter(selectedShapes, slideWidth, slideHeight);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }

        private void DefaultAlignLeftButton_Click(object sender, RoutedEventArgs e)
        {
            DefaultAlign(Office.MsoAlignCmd.msoAlignLefts);
        }

        private void DefaultAlignRightButton_Click(object sender, RoutedEventArgs e)
        {
            DefaultAlign(Office.MsoAlignCmd.msoAlignRights);
        }

        private void DefaultAlignTopButton_Click(object sender, RoutedEventArgs e)
        {
            DefaultAlign(Office.MsoAlignCmd.msoAlignTops);
        }

        private void DefaultAlignBottomButton_Click(object sender, RoutedEventArgs e)
        {
            DefaultAlign(Office.MsoAlignCmd.msoAlignBottoms);
        }

        private void DefaultAlignCenterButton_Click(object sender, RoutedEventArgs e)
        {
            DefaultAlign(Office.MsoAlignCmd.msoAlignCenters);
        }

        private void DefaultAlignMiddleButton_Click(object sender, RoutedEventArgs e)
        {
            DefaultAlign(Office.MsoAlignCmd.msoAlignMiddles);
        }

        private void DefaultAlign(Office.MsoAlignCmd direction)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = this.GetCurrentSelection().ShapeRange;
                PositionsLabMain.DefaultAlign(selectedShapes, direction);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }
        #endregion

        #region Adjoin
        private void AdjoinHorizontalButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                PositionsLabMain.AdjoinHorizontal(selectedShapes);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }

        private void AdjoinVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                PositionsLabMain.AdjoinVertical(selectedShapes);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }
        #endregion

        #region Distribute
        private void DistributeHorizontalButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                var slideWidth = this.GetCurrentPresentation().SlideWidth;
                PositionsLabMain.DistributeHorizontal(selectedShapes, slideWidth);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }

        private void DistributeVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                var slideHeight = this.GetCurrentPresentation().SlideHeight;
                PositionsLabMain.DistributeVertical(selectedShapes, slideHeight);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }

        private void DistributeCenterButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                var slideWidth = this.GetCurrentPresentation().SlideWidth;
                var slideHeight = this.GetCurrentPresentation().SlideHeight;
                PositionsLabMain.DistributeCenter(selectedShapes, slideWidth, slideHeight);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }

        // TODO: Remove after confirming that Distribute Shapes shall be part of settings
        private void DistributeShapesButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                PositionsLabMain.DistributeShapes(selectedShapes);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }

        private void DistributeGridButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                var numShapesSelected = selectedShapes.Count;
                var rowLength = (int)Math.Ceiling(Math.Sqrt(numShapesSelected));
                var colLength = (int)Math.Ceiling((double)numShapesSelected / rowLength);

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
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }  
        }
        #endregion

        #region Reorder
        private void SwapPositionsButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                PositionsLabMain.Swap(selectedShapes);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }
        #endregion

        #region Adjustment
        private void RotationButton_Click(object sender, RoutedEventArgs e)
        {
            var noShapesSelected = this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes;

            if (noShapesSelected)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            var selectedShapes = this.GetCurrentSelection().ShapeRange;

            if (selectedShapes.Count <= 1)
            {
                ShowErrorMessageBox(ErrorMessageFewerThanTwoSelection);
                return;
            }

            if (_isLockAxisMode)
            {
                DisableLockAxisMode();
            }

            _isRotationMode = true;

            var currentSlide = this.GetCurrentSlide();

            _refPoint = selectedShapes[1];
            _shapesToBeRotated = ConvertShapeRangeToList_Legacy(selectedShapes, 2);
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
            var p = System.Windows.Forms.Control.MousePosition;

            var prevAngle = (float)PositionsLabMain.AngleBetweenTwoPoints(ConvertSlidePointToScreenPoint(Graphics.GetCenterPoint(_refPoint)), _prevMousePos);
            var angle = (float)PositionsLabMain.AngleBetweenTwoPoints(ConvertSlidePointToScreenPoint(Graphics.GetCenterPoint(_refPoint)), p) - prevAngle;
            var origin = Graphics.GetCenterPoint(_refPoint);

            foreach (var currentShape in _shapesToBeRotated)
            {
                var unrotatedCenter = Graphics.GetCenterPoint(currentShape);
                var rotatedCenter = Graphics.RotatePoint(unrotatedCenter, origin, angle);

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
                var p = System.Windows.Forms.Control.MousePosition;
                var selectedShape = GetShapeDirectlyBelowMousePos(_allShapesInSlide, p);

                if (selectedShape == null)
                {
                    DisableRotationMode();

                    if (_isLockAxisMode)
                    {
                        StartLockAxisMode();
                    }
                    return;
                }

                var isShapeToBeRotated = _shapesToBeRotated.Contains(selectedShape);
                var isRefPoint = _refPoint.Id == selectedShape.Id;

                if (!isShapeToBeRotated && !isRefPoint)
                {
                    DisableRotationMode();

                    if (_isLockAxisMode)
                    {
                        StartLockAxisMode();
                    }
                    return;
                }

                this.StartNewUndoEntry();

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

            var noShapesSelected = this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes;

            if (_shapesToBeMoved == null && noShapesSelected)
            {
                return;
            }

            if (_shapesToBeMoved == null)
            {
                _shapesToBeMoved = this.GetCurrentSelection().ShapeRange;
                _initialPos = new float[_shapesToBeMoved.Count, 2];
                for (var i = 0; i < _shapesToBeMoved.Count; i++)
                {
                    var s = _shapesToBeMoved[i + 1];
                    _initialPos[i, Left] = s.Left;
                    _initialPos[i, Top] = s.Top;
                }
            }

            //Remove dragging control of user
            this.GetCurrentSelection().Unselect();

            var currentMousePos = System.Windows.Forms.Control.MousePosition;

            float diffX = currentMousePos.X - _initialMousePos.X;
            float diffY = currentMousePos.Y - _initialMousePos.Y;

            for (var i = 0; i < _shapesToBeMoved.Count; i++)
            {
                var s = _shapesToBeMoved[i + 1];
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
                var p = System.Windows.Forms.Control.MousePosition;
                var currentSlide = this.GetCurrentSlide();
                _allShapesInSlide = ConvertShapesToList(currentSlide.Shapes);
                var selectedShape = GetShapeDirectlyBelowMousePos(_allShapesInSlide, p);

                if (selectedShape == null || PPKeyboard.IsCtrlPressed() || PPKeyboard.IsShiftPressed())
                {
                    return;
                }

                this.StartNewUndoEntry();

                _initialMousePos = p;
                _dispatcherTimer.Start();
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "LockAxis");
            }
        }

        #endregion

        #region Snap
        private void SnapHorizontalButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList_Legacy(this.GetCurrentSelection().ShapeRange, 1);
                PositionsLabMain.SnapHorizontal(selectedShapes);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }

        private void SnapVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList_Legacy(this.GetCurrentSelection().ShapeRange, 1);
                PositionsLabMain.SnapVertical(selectedShapes);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }

        private void SnapAwayButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            try
            {
                if (_previewIsExecuted)
                {
                    UndoPreview();
                }
                this.StartNewUndoEntry();
                var selectedShapes = ConvertShapeRangeToList_Legacy(this.GetCurrentSelection().ShapeRange, 1);
                PositionsLabMain.SnapAway(selectedShapes);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
        }
        #endregion
        #endregion

        #region Preview Behaviour
        private void AlignLeftButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            { 
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                this.StartNewUndoEntry();
                PositionsLabMain.AlignLeft(selectedShapes);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }

        private void AlignRightButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            {
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                this.StartNewUndoEntry();
                var slideWidth = this.GetCurrentPresentation().SlideWidth;
                PositionsLabMain.AlignRight(selectedShapes, slideWidth);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }

        private void AlignTopButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            {
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                this.StartNewUndoEntry();
                PositionsLabMain.AlignTop(selectedShapes);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }

        private void AlignBottomButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            {
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                this.StartNewUndoEntry();
                var slideHeight = this.GetCurrentPresentation().SlideHeight;
                PositionsLabMain.AlignBottom(selectedShapes, slideHeight);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }

        private void AlignMiddleButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            {
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                this.StartNewUndoEntry();
                var slideHeight = this.GetCurrentPresentation().SlideHeight;
                PositionsLabMain.AlignMiddle(selectedShapes, slideHeight);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }

        private void AlignCenterButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            {
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                this.StartNewUndoEntry();
                var slideHeight = this.GetCurrentPresentation().SlideHeight;
                var slideWidth = this.GetCurrentPresentation().SlideWidth;
                PositionsLabMain.AlignCenter(selectedShapes, slideHeight, slideWidth);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }

        private void DefaultAlignCenterButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            DefaultAlignPreview(Office.MsoAlignCmd.msoAlignCenters);
        }

        private void DefaultAlignMiddleButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            DefaultAlignPreview(Office.MsoAlignCmd.msoAlignMiddles);
        }

        private void DefaultAlignLeftButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            DefaultAlignPreview(Office.MsoAlignCmd.msoAlignLefts);
        }

        private void DefaultAlignRightButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            DefaultAlignPreview(Office.MsoAlignCmd.msoAlignRights);
        }

        private void DefaultAlignTopButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            DefaultAlignPreview(Office.MsoAlignCmd.msoAlignTops);
        }

        private void DefaultAlignBottomButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            DefaultAlignPreview(Office.MsoAlignCmd.msoAlignBottoms);
        }

        private void DefaultAlignPreview(Office.MsoAlignCmd direction)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            {
                var selectedShapes = this.GetCurrentSelection().ShapeRange;
                this.StartNewUndoEntry();
                PositionsLabMain.DefaultAlign(selectedShapes, direction);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }

        private void AdjoinHorizontalButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            {
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                this.StartNewUndoEntry();
                PositionsLabMain.AdjoinHorizontal(selectedShapes);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }

        private void AdjoinVerticalButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            {
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                this.StartNewUndoEntry();
                PositionsLabMain.AdjoinVertical(selectedShapes);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }

        private void DistributeHorizontalButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            {
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                this.StartNewUndoEntry();
                var slideWidth = this.GetCurrentPresentation().SlideWidth;
                PositionsLabMain.DistributeHorizontal(selectedShapes, slideWidth);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }

        private void DistributeVerticalButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            {
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                this.StartNewUndoEntry();
                var slideHeight = this.GetCurrentPresentation().SlideHeight;
                PositionsLabMain.DistributeVertical(selectedShapes, slideHeight);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }

        private void DistributeCenterButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            {
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                this.StartNewUndoEntry();
                var slideWidth = this.GetCurrentPresentation().SlideWidth;
                var slideHeight = this.GetCurrentPresentation().SlideHeight;
                PositionsLabMain.DistributeCenter(selectedShapes, slideWidth, slideHeight);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }

        private void DistributeShapesButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            {
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                this.StartNewUndoEntry();
                PositionsLabMain.DistributeShapes(selectedShapes);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }

        private void SwapPositionsButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            {
                var selectedShapes = ConvertShapeRangeToList(this.GetCurrentSelection().ShapeRange, 1);
                this.StartNewUndoEntry();
                PositionsLabMain.Swap(selectedShapes);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }

        private void SnapHorizontalButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            {
                var selectedShapes = ConvertShapeRangeToList_Legacy(this.GetCurrentSelection().ShapeRange, 1);
                this.StartNewUndoEntry();
                PositionsLabMain.SnapHorizontal(selectedShapes);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }

        private void SnapVerticalButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            {
                var selectedShapes = ConvertShapeRangeToList_Legacy(this.GetCurrentSelection().ShapeRange, 1);
                this.StartNewUndoEntry();
                PositionsLabMain.SnapVertical(selectedShapes);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }

        private void SnapAwayButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }
            try
            {
                var selectedShapes = ConvertShapeRangeToList_Legacy(this.GetCurrentSelection().ShapeRange, 1);
                this.StartNewUndoEntry();
                PositionsLabMain.SnapAway(selectedShapes);
                _previewIsExecuted = true;
            }
            catch
            {
                return;
            }
        }
        #endregion

        #region Helper
        private Shape AddReferencePoint(PowerPoint.Shapes shapes, float left, float top)
        {
            return shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left, top, RefpointRadius, RefpointRadius);
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
            var epsilon = 0.00001f;

            var centerPoint = ConvertSlidePointToScreenPoint(Graphics.GetCenterPoint(shape));
            var rotatedMousePos = Graphics.RotatePoint(p, centerPoint, -shape.Rotation);

            var x1 = PointsToScreenPixelsX(shape.Left);
            var y1 = PointsToScreenPixelsY(shape.Top);
            var x2 = PointsToScreenPixelsX(shape.Left + shape.Width);
            var y2 = PointsToScreenPixelsY(shape.Top + shape.Height);

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

            foreach (var s in shapes)
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

        private List<PPShape> ConvertShapeRangeToList (PowerPoint.ShapeRange range, int index)
        {
            var shapes = new List<PPShape>();

            for (var i = index; i <= range.Count; i++)
            {
                shapes.Add(new PPShape(range[i]));
            }

            return shapes;
        }

        private List<Shape> ConvertShapeRangeToList_Legacy(PowerPoint.ShapeRange range, int index)
        {
            var shapes = new List<Shape>();

            for (var i = index; i <= range.Count; i++)
            {
                shapes.Add(range[i]);
            }

            return shapes;
        }

        private List<Shape> ConvertShapesToList(PowerPoint.Shapes shapes)
        {
            var listOfShapes = new List<Shape>();

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
            foreach (var s in shapes)
            {
                s.Select(Office.MsoTriState.msoFalse);
            }
        }

        #endregion

        #region Settings
        private void AlignSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            if (_alignSettingsDialog == null || !_alignSettingsDialog.IsOpen)
            {
                _alignSettingsDialog = new AlignSettingsDialog();
                _alignSettingsDialog.Show();
            }
            else
            {
                _alignSettingsDialog.Activate();
            }
        }

        private void DistributeSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            if (_distributeSettingsDialog == null || !_distributeSettingsDialog.IsOpen)
            {
                _distributeSettingsDialog = new DistributeSettingsDialog();
                _distributeSettingsDialog.Show();
            }
            else
            {
                _distributeSettingsDialog.Activate();
            }
        }

        private void ReorderSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            if (_reorderSettingsDialog == null || !_reorderSettingsDialog.IsOpen)
            {
                _reorderSettingsDialog = new ReorderSettingsDialog();
                _reorderSettingsDialog.Show();
            }
            else
            {
                _reorderSettingsDialog.Activate();
            }
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

        #region Error Handling
        public void ShowErrorMessageBox(string content, Exception exception = null)
        {

            if (exception == null)
            {
                MessageBox.Show(content, "Error");
                return;
            }
            
            var errorMessage = GetErrorMessage(exception.Message);
            if (!string.Equals(errorMessage, ErrorMessageUndefined, StringComparison.Ordinal))
            {
                MessageBox.Show(content, "Error");
            }
            else
            {
                Views.ErrorDialogWrapper.ShowDialog("Error", content, exception);
            }
        }

        private string GetErrorMessage(string errorMsg)
        {
            switch (errorMsg)
            {
                case ErrorMessageNoSelection:
                    return ErrorMessageNoSelection;
                case ErrorMessageFewerThanTwoSelection:
                    return ErrorMessageFewerThanTwoSelection;
                case ErrorMessageFewerThanThreeSelection:
                    return ErrorMessageFewerThanThreeSelection;
                default:
                    return ErrorMessageUndefined;
            }
        }

        private void IgnoreExceptionThrown() { }

        #endregion

        private void UndoPreview(object sender, System.Windows.Input.MouseEventArgs e)
        {
            UndoPreview();
        }

        private void UndoPreview()
        {
            if (_previewIsExecuted)
            {
                this.ExecuteOfficeCommand("Undo");
                _previewIsExecuted = false;
                GC.Collect();
            }
        }
    }
}
