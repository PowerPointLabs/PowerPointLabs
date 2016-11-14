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
using System.Windows.Input;
using System.Windows.Controls.Primitives;
using Media = System.Windows.Media;

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
        private const string ErrorMessageFewerThanThreeSelection = TextCollection.PositionsLabText.ErrorFewerThanThreeSelection;
        private const string ErrorMessageFewerThanFourSelection = TextCollection.PositionsLabText.ErrorFewerThanFourSelection;
        private const string ErrorMessageFunctionNotSupportedForExtremeShapes = 
            TextCollection.PositionsLabText.ErrorFunctionNotSupportedForWithinShapes;
        private const string ErrorMessageFunctionNotSupportedForSlide =
            TextCollection.PositionsLabText.ErrorFunctionNotSupportedForSlide;
        private const string ErrorMessageFunctionNotSuppertedForOverlapRefShapeCenter =
            TextCollection.PositionsLabText.ErrorFunctionNotSupportedForOverlapRefShapeCenter;
        private const string ErrorMessageUndefined = TextCollection.PositionsLabText.ErrorUndefined;

        //Variable for preview
        private bool _previewIsExecuted = false;
        private delegate void PreviewCallBack();
        private PreviewCallBack _previewCallBack;
        private static Dictionary<int, PositionShapeProperties> allShapePos = new Dictionary<int, PositionShapeProperties>();

        //Variables for lock axis
        private const int Left = 0;
        private const int Top = 1;
        private static List<Shape> _shapesToBeMoved;
        private static System.Drawing.Point _initialMousePos;
        private float[,] _initialPos;

        //Variables for rotation
        private const float RefpointRadius = 10;
        private static Shape _refPoint;
        private static List<Shape> _shapesToBeRotated = new List<Shape>();
        private static List<Shape> _allShapesInSlide = new List<Shape>();
        private static System.Drawing.Point _prevMousePos;
        private static PowerPoint.ShapeRange _selectedRange;

        //Variables for key binding
        private const int CtrlProportion = 5;

        //Variables for settings
        private AlignSettingsDialog _alignSettingsDialog;
        private DistributeSettingsDialog _distributeSettingsDialog;
        private ReorderSettingsDialog _reorderSettingsDialog;
        private ReorientSettingsDialog _reorientSettingsDialog;

        public PositionsPaneWpf()
        {
            PositionsLabMain.InitPositionsLab();
            InitializeComponent();
            InitializeHotKeys();
            _dispatcherTimer.Interval = TimeSpan.FromMilliseconds(10);
            Focusable = true;
        }
        
        private void InitializeHotKeys()
        {
            var buttonActionMapping = new Dictionary<ToggleButton, Action<bool, bool>>();
            buttonActionMapping.Add(rotationButton, RotateSlightly);
            buttonActionMapping.Add(duplicateRotationButton, RotateSlightly);

            Action<Native.VirtualKey, bool> bindHotKeys =
                (key, direction) =>
                {
                    PPKeyboard.AddKeydownAction(key, RunOnlyWhenActivated(buttonActionMapping, direction, true));
                    PPKeyboard.AddKeydownAction(key, RunOnlyWhenActivated(buttonActionMapping, direction, false), ctrl: true);
                };

            bindHotKeys(Native.VirtualKey.VK_LEFT, false);
            bindHotKeys(Native.VirtualKey.VK_UP, false);
            bindHotKeys(Native.VirtualKey.VK_RIGHT, true);
            bindHotKeys(Native.VirtualKey.VK_DOWN, true);
        }

        private Func<bool> RunOnlyWhenActivated(Dictionary<ToggleButton, Action<bool, bool>> buttonActionMapping, bool direction, bool isLarge)
        {
            return () =>
            {
                var positionsPane = this.GetTaskPane(typeof(PositionsPane));
                if (positionsPane == null || !positionsPane.Visible)
                {
                    return false;
                }

                foreach (var mapping in buttonActionMapping)
                {
                    var button = mapping.Key;
                    var action = mapping.Value;

                    if ((bool)button.IsChecked)
                    {
                        action(direction, isLarge);
                        return true;
                    }
                }

                return false;
            };
        }
        
        private void RotateSlightly(bool isClockwise, bool isLarge)
        {
            var origin = Graphics.GetCenterPoint(_refPoint);
            var angle = (isClockwise) ? 1f : -1f;
            if (isLarge)
            {
                angle *= CtrlProportion;
            }

            foreach (var currentShape in _shapesToBeRotated)
            {
                PositionsLabMain.Rotate(currentShape, origin, angle, PositionsLabMain.ReorientShapeOrientation);
            }
        }

        #region Click Behaviour
        #region Align
        private void AlignLeftButton_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> positionsAction = shapes => PositionsLabMain.AlignLeft(shapes);
            ExecutePositionsAction(positionsAction, false);
        }

        private void AlignRightButton_Click(object sender, RoutedEventArgs e)
        {
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            Action<PowerPoint.ShapeRange, float> positionsAction = (shapes, width) => PositionsLabMain.AlignRight(shapes, width);
            ExecutePositionsAction(positionsAction, slideWidth, false);
        }

        private void AlignTopButton_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> positionsAction = shapes => PositionsLabMain.AlignTop(shapes);
            ExecutePositionsAction(positionsAction, false);
        }

        private void AlignBottomButton_Click(object sender, RoutedEventArgs e)
        {
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<PowerPoint.ShapeRange, float> positionsAction = (shapes, height) => PositionsLabMain.AlignBottom(shapes, height);
            ExecutePositionsAction(positionsAction, slideHeight, false);
        }

        private void AlignHorizontalCenterButton_Click(object sender, RoutedEventArgs e)
        {
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<PowerPoint.ShapeRange, float> positionsAction = (shapes, height) => PositionsLabMain.AlignHorizontalCenter(shapes, height);
            ExecutePositionsAction(positionsAction, slideHeight, false);
        }

        private void AlignVerticalCenterButton_Click(object sender, RoutedEventArgs e)
        {
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            Action<PowerPoint.ShapeRange, float> positionsAction = (shapes, width) => PositionsLabMain.AlignVerticalCenter(shapes, width);
            ExecutePositionsAction(positionsAction, slideWidth, false);
        }

        private void AlignCenterButton_Click(object sender, RoutedEventArgs e)
        {
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            Action<PowerPoint.ShapeRange, float, float> positionsAction = (shapes, height, width) => PositionsLabMain.AlignCenter(shapes, height, width);
            ExecutePositionsAction(positionsAction, slideHeight, slideWidth, false);
        }

        private void AlignRadialButton_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.AlignRadial(shapes);
            ExecutePositionsAction(positionsAction, false, isConvertPPShape: false);
        }
        #endregion

        #region Adjoin
        private void AdjoinHorizontalButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AdjoinWithoutAligning();
            Action<List<PPShape>> positionsAction = (shapes) => PositionsLabMain.AdjoinHorizontal(shapes);
            ExecutePositionsAction(positionsAction, false);
        }
        private void AdjoinHorizontalWithAlignButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AdjoinWithAligning();
            Action<List<PPShape>> positionsAction = (shapes) => PositionsLabMain.AdjoinHorizontal(shapes);
            ExecutePositionsAction(positionsAction, false);
        }

        private void AdjoinVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AdjoinWithoutAligning();
            Action<List<PPShape>> positionsAction = (shapes) => PositionsLabMain.AdjoinVertical(shapes);
            ExecutePositionsAction(positionsAction, false);
        }

        private void AdjoinVerticalWithAlignButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabMain.AdjoinWithAligning();
            Action<List<PPShape>> positionsAction = (shapes) => PositionsLabMain.AdjoinVertical(shapes);
            ExecutePositionsAction(positionsAction, false);
        }
        #endregion

        #region Distribute
        private void DistributeHorizontalButton_Click(object sender, RoutedEventArgs e)
        {
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            Action<List<PPShape>, float> positionsAction = (shapes, width) => PositionsLabMain.DistributeHorizontal(shapes, width);
            ExecutePositionsAction(positionsAction, slideWidth, false);
        }

        private void DistributeVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<List<PPShape>, float> positionsAction = (shapes, height) => PositionsLabMain.DistributeVertical(shapes, height);
            ExecutePositionsAction(positionsAction, slideHeight, false);
        }

        private void DistributeCenterButton_Click(object sender, RoutedEventArgs e)
        {
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<List<PPShape>, float, float> positionsAction = (shapes, width, height) => PositionsLabMain.DistributeCenter(shapes, width, height);
            ExecutePositionsAction(positionsAction, slideWidth, slideHeight, false);
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
                var selectedShapes = this.GetCurrentSelection().ShapeRange;
                var numShapesSelected = selectedShapes.Count;
                var rowLength = (int)Math.Ceiling(Math.Sqrt(numShapesSelected));
                var colLength = (int)Math.Ceiling((double)numShapesSelected / rowLength);

                if (_positionsDistributeGridDialog == null || !_positionsDistributeGridDialog.IsOpen)
                {
                    _positionsDistributeGridDialog = new PositionsDistributeGridDialog(selectedShapes, rowLength, colLength);
                    _positionsDistributeGridDialog.ShowDialog();
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

        private void DistributeRadialButton_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, false, isConvertPPShape: false);
        }

        #endregion

        #region Reorder
        private void SwapPositionsButton_Click(object sender, RoutedEventArgs e)
        {
            Action<List<PPShape>, bool> positionsAction = (shapes, isPreview) => PositionsLabMain.Swap(shapes, isPreview);
            ExecutePositionsAction(positionsAction, false, false);
        }
        #endregion

        #region Adjustment

        private void RotationButton_Click(object sender, RoutedEventArgs e)
        {
            var noShapesSelected = this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes;
            var button = (ToggleButton)sender;

            if (noShapesSelected)
            {
                button.IsChecked = false;
                ShowErrorMessageBox(ErrorMessageFewerThanTwoSelection);
                return;
            }

            var selectedShapes = this.GetCurrentSelection().ShapeRange;

            if (selectedShapes.Count <= 1)
            {
                button.IsChecked = false;
                ShowErrorMessageBox(ErrorMessageFewerThanTwoSelection);
                return;
            }

            ClearAllEventHandlers();

            var currentSlide = this.GetCurrentSlide();

            _refPoint = selectedShapes[1];
            _shapesToBeRotated = ConvertShapeRangeToShapeList(selectedShapes, 2);
            _allShapesInSlide = ConvertShapesToShapeList(currentSlide.Shapes);
            _selectedRange = selectedShapes;

            StartRotationMode();

            // for key binding to work when select shapes first, then open panel and click button
            PPKeyboard.SetSlideViewWindowFocused();
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
                PositionsLabMain.Rotate(currentShape, origin, angle, PositionsLabMain.ReorientShapeOrientation);
            }

            _prevMousePos = p;
        }

        void _leftMouseUpListener_Rotation(object sender, SysMouseEventInfo e)
        {
            _dispatcherTimer.Stop();
            _selectedRange.Select();
        }

        void _leftMouseDownListener_Rotation(object sender, SysMouseEventInfo e)
        {
            try
            {
                var button = ((bool)rotationButton.IsChecked) ? rotationButton :
                             ((bool)duplicateRotationButton.IsChecked) ? duplicateRotationButton
                                                                       : null;

                if (button.IsMouseOver)
                {
                    DisableRotationMode();
                    return;
                }

                var p = System.Windows.Forms.Control.MousePosition;
                var selectedShape = GetShapeDirectlyBelowMousePos(_allShapesInSlide, p);

                if (selectedShape == null)
                {
                    DisableRotationMode();
                    button.IsChecked = false;
                    return;
                }

                var isShapeToBeRotated = _shapesToBeRotated.Contains(selectedShape);
                var isRefPoint = _refPoint.Id == selectedShape.Id;

                if (!isShapeToBeRotated && !isRefPoint)
                {
                    DisableRotationMode();
                    button.IsChecked = false;
                    return;
                }

                this.StartNewUndoEntry();

                if (isRefPoint)
                {
                    this.GetCurrentSelection().Unselect();
                    return;
                }

                if (button == duplicateRotationButton)
                {
                    foreach (var currentShape in _shapesToBeRotated)
                    {
                        var duplicatedShape = currentShape.Duplicate()[1];
                        duplicatedShape.Left -= 12;
                        duplicatedShape.Top -= 12;
                        Graphics.MoveZToJustBehind(duplicatedShape, currentShape);
                    }
                }

                _prevMousePos = p;
                _dispatcherTimer.Start();
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "Rotation");
            }
        }

        private void LockAxisButton_Click(object sender, RoutedEventArgs e)
        {
            var noShapesSelected = this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes;

            if (noShapesSelected)
            {
                lockAxisButton.IsChecked = false;
                ShowErrorMessageBox(ErrorMessageNoSelection);
                return;
            }

            var selectedShapes = this.GetCurrentSelection().ShapeRange;

            ClearAllEventHandlers();

            var currentSlide = this.GetCurrentSlide();

            _shapesToBeMoved = ConvertShapeRangeToShapeList(selectedShapes, 1);
            _allShapesInSlide = ConvertShapesToShapeList(currentSlide.Shapes);
            _selectedRange = selectedShapes;

            StartLockAxisMode();
        }

        private void LockAxisHandler(object sender, EventArgs e)
        {
            //Remove dragging control of user
            this.GetCurrentSelection().Unselect();

            var currentMousePos = System.Windows.Forms.Control.MousePosition;

            float diffX = currentMousePos.X - _initialMousePos.X;
            float diffY = currentMousePos.Y - _initialMousePos.Y;

            for (var i = 0; i < _shapesToBeMoved.Count; i++)
            {
                var s = _shapesToBeMoved[i];
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
            _selectedRange.Select();
        }

        void _leftMouseDownListener_LockAxis(object sender, SysMouseEventInfo e)
        {
            try
            {
                if (lockAxisButton.IsMouseOver)
                {
                    DisableLockAxisMode();
                    return;
                }

                var p = System.Windows.Forms.Control.MousePosition;
                var currentSlide = this.GetCurrentSlide();
                var selectedShape = GetShapeDirectlyBelowMousePos(_allShapesInSlide, p);

                if (selectedShape == null)
                {
                    DisableLockAxisMode();
                    lockAxisButton.IsChecked = false;
                    return;
                }

                var isShapeToBeMoved = _shapesToBeMoved.Contains(selectedShape);

                if (!isShapeToBeMoved)
                {
                    DisableLockAxisMode();
                    lockAxisButton.IsChecked = false;
                    return;
                }

                this.StartNewUndoEntry();

                _initialPos = new float[_shapesToBeMoved.Count, 2];
                for (var i = 0; i < _shapesToBeMoved.Count; i++)
                {
                    var s = _shapesToBeMoved[i];
                    _initialPos[i, Left] = s.Left;
                    _initialPos[i, Top] = s.Top;
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

        #region Snap
        private void SnapHorizontalButton_Click(object sender, RoutedEventArgs e)
        {
            Action<List<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapHorizontal(shapes);
            ExecutePositionsAction(positionsAction, false);
        }

        private void SnapVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            Action<List<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapVertical(shapes);
            ExecutePositionsAction(positionsAction, false);
        }

        private void SnapAwayButton_Click(object sender, RoutedEventArgs e)
        {
            Action<List<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapAway(shapes);
            ExecutePositionsAction(positionsAction, false);
        }
        #endregion
        #endregion

        #region Preview Behaviour
        private void AlignLeftButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange> positionsAction = shapes => PositionsLabMain.AlignLeft(shapes);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, true);
            };
            PreviewHandler();
        }

        private void AlignRightButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            Action<PowerPoint.ShapeRange, float> positionsAction = (shapes, width) => PositionsLabMain.AlignRight(shapes, width);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, slideWidth, true);
            };
            PreviewHandler();
        }

        private void AlignTopButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange> positionsAction = shapes => PositionsLabMain.AlignTop(shapes);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, true);
            };
            PreviewHandler();
        }

        private void AlignBottomButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<PowerPoint.ShapeRange, float> positionsAction = (shapes, height) => PositionsLabMain.AlignBottom(shapes, height);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, slideHeight, true);
            };
            PreviewHandler();
        }

        private void AlignHorizontalCenterButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<PowerPoint.ShapeRange, float> positionsAction = (shapes, height) => PositionsLabMain.AlignHorizontalCenter(shapes, height);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, slideHeight, true);
            };
            PreviewHandler();
        }

        private void AlignVerticalCenterButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            Action<PowerPoint.ShapeRange, float> positionsAction = (shapes, width) => PositionsLabMain.AlignVerticalCenter(shapes, width);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, slideWidth, true);
            };
            PreviewHandler();
        }

        private void AlignCenterButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            Action<PowerPoint.ShapeRange, float, float> positionsAction = (shapes, height, width) => PositionsLabMain.AlignCenter(shapes, height, width);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, slideHeight, slideWidth, true);
            };
            PreviewHandler();
        }

        private void AlignRadialButton_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.AlignRadial(shapes);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, true, isConvertPPShape: false);
            };
            PreviewHandler();
        }

        private void AdjoinHorizontalButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            PositionsLabMain.AdjoinWithoutAligning();
            Action<List<PPShape>> positionsAction = (shapes) => PositionsLabMain.AdjoinHorizontal(shapes);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, true);
            };
            PreviewHandler();
        }

        private void AdjoinHorizontalWithAlignButton_MouseEnter(object sender, MouseEventArgs e)
        {
            PositionsLabMain.AdjoinWithAligning();
            Action<List<PPShape>> positionsAction = (shapes) => PositionsLabMain.AdjoinHorizontal(shapes);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, true);
            };
            PreviewHandler();
        }

        private void AdjoinVerticalButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            PositionsLabMain.AdjoinWithoutAligning();
            Action<List<PPShape>> positionsAction = (shapes) => PositionsLabMain.AdjoinVertical(shapes);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, true);
            };
            PreviewHandler();
        }

        private void AdjoinVerticalWithAlignButton_MouseEnter(object sender, MouseEventArgs e)
        {
            PositionsLabMain.AdjoinWithAligning();
            Action<List<PPShape>> positionsAction = (shapes) => PositionsLabMain.AdjoinVertical(shapes);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, true);
            };
            PreviewHandler();
        }

        private void DistributeHorizontalButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            Action<List<PPShape>, float> positionsAction = (shapes, width) => PositionsLabMain.DistributeHorizontal(shapes, width);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, slideWidth, true);
            };
            PreviewHandler();
        }

        private void DistributeVerticalButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<List<PPShape>, float> positionsAction = (shapes, height) => PositionsLabMain.DistributeVertical(shapes, height);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, slideHeight, true);
            };
            PreviewHandler();
        }

        private void DistributeCenterButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<List<PPShape>, float, float> positionsAction = (shapes, width, height) => PositionsLabMain.DistributeCenter(shapes, width, height);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, slideWidth, slideHeight, true);
            };
            PreviewHandler();
        }

        private void DistributeRadialButton_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, true, isConvertPPShape: false);
            };
            PreviewHandler();
        }

        private void SwapPositionsButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            Action<List<PPShape>, bool> positionsAction = (shapes, isPreview) => PositionsLabMain.Swap(shapes, isPreview);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, true, true);
            };
            PreviewHandler();
        }

        private void SnapHorizontalButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            Action<List<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapHorizontal(shapes);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, true);
            };
            PreviewHandler();
        }

        private void SnapVerticalButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            Action<List<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapVertical(shapes);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, true);
            };
            PreviewHandler();
        }

        private void SnapAwayButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            Action<List<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapAway(shapes);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, true);
            };
            PreviewHandler();
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

        private List<PPShape> ConvertShapeRangeToPPShapeList (PowerPoint.ShapeRange range, int index)
        {
            var shapes = new List<PPShape>();

            for (var i = index; i <= range.Count; i++)
            {
                var s = range[i];
                if (s.Type.Equals(Office.MsoShapeType.msoPicture))
                {
                    shapes.Add(new PPShape(range[i], false));
                }
                else
                {
                    shapes.Add(new PPShape(range[i]));
                }
            }

            return shapes;
        }

        private List<Shape> ConvertShapeRangeToShapeList(PowerPoint.ShapeRange range, int index)
        {
            var shapes = new List<Shape>();

            for (var i = index; i <= range.Count; i++)
            {
                shapes.Add(range[i]);
            }

            return shapes;
        }

        private List<Shape> ConvertShapesToShapeList(PowerPoint.Shapes shapes)
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
                _alignSettingsDialog.ShowDialog();
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
                _distributeSettingsDialog.ShowDialog();
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
                _reorderSettingsDialog.ShowDialog();
            }
            else
            {
                _reorderSettingsDialog.Activate();
            }
        }

        private void ReorientSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            if (_reorientSettingsDialog == null || !_reorientSettingsDialog.IsOpen)
            {
                _reorientSettingsDialog = new ReorientSettingsDialog();
                _reorientSettingsDialog.ShowDialog();
            }
            else
            {
                _reorientSettingsDialog.Activate();
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

        private void StartRotationMode()
        {
            _dispatcherTimer.Tick += RotationHandler;

            _leftMouseUpListener = new LMouseUpListener();
            _leftMouseUpListener.LButtonUpClicked += _leftMouseUpListener_Rotation;

            _leftMouseDownListener = new LMouseDownListener();
            _leftMouseDownListener.LButtonDownClicked += _leftMouseDownListener_Rotation;
        }

        private void DisableRotationMode()
        {
            ClearAllEventHandlers();
            _selectedRange = null;
            _refPoint = null;
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

        private void DisableLockAxisMode()
        {
            ClearAllEventHandlers();
            _selectedRange = null;
            _shapesToBeMoved = null;
            _initialMousePos = new System.Drawing.Point();
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
                case ErrorMessageFewerThanFourSelection:
                    return ErrorMessageFewerThanFourSelection;
                case ErrorMessageFunctionNotSupportedForExtremeShapes:
                    return ErrorMessageFunctionNotSupportedForExtremeShapes;
                case ErrorMessageFunctionNotSupportedForSlide:
                    return ErrorMessageFunctionNotSupportedForSlide;
                case ErrorMessageFunctionNotSuppertedForOverlapRefShapeCenter:
                    return ErrorMessageFunctionNotSuppertedForOverlapRefShapeCenter;
                default:
                    return ErrorMessageUndefined;
            }
        }

        private void IgnoreExceptionThrown() { }

        #endregion

        #region Helper
        // align left and top
        public void ExecutePositionsAction(Action<PowerPoint.ShapeRange> positionsAction, bool isPreview, bool isConvertPPShape = true)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(ErrorMessageNoSelection);
                }
                return;
            }

            PowerPoint.ShapeRange simulatedShapes = null;

            try
            {
                var selectedShapes = this.GetCurrentSelection().ShapeRange;

                if (isPreview)
                {
                    SaveSelectedShapePositions(selectedShapes, allShapePos);
                }
                else
                {
                    UndoPreview();
                    _previewCallBack = null;
                    this.StartNewUndoEntry();
                }

                simulatedShapes = DuplicateShapes(selectedShapes);

                if (PositionsLabMain.AlignReference == PositionsLabMain.AlignReferenceObject.PowerpointDefaults)
                {
                    positionsAction.Invoke(selectedShapes);
                }
                else if (isConvertPPShape)
                {
                    var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                    var initialPositions = SaveOriginalPositions(simulatedPPShapes);

                    positionsAction.Invoke(simulatedShapes);

                    SyncShapes(selectedShapes, simulatedShapes, initialPositions);
                }
                else
                {
                    positionsAction.Invoke(simulatedShapes);

                    SyncShapes(selectedShapes, simulatedShapes);
                }
                if (isPreview)
                {
                    _previewIsExecuted = true;
                }
            }
            catch (Exception ex)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(ex.Message, ex);
                }
            }
            finally
            {
                if (simulatedShapes != null)
                {
                    simulatedShapes.Delete();
                    GC.Collect();
                }
            }
        }

        // Align right, bottom, vertical center, horizontal center
        public void ExecutePositionsAction(Action<PowerPoint.ShapeRange, float> positionsAction, float dimension, bool isPreview)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(ErrorMessageNoSelection);
                }
                return;
            }

            PowerPoint.ShapeRange simulatedShapes = null;

            try
            {
                var selectedShapes = this.GetCurrentSelection().ShapeRange;

                if (isPreview)
                {
                    SaveSelectedShapePositions(selectedShapes, allShapePos);
                }
                else
                {
                    UndoPreview();
                    _previewCallBack = null;
                    this.StartNewUndoEntry();
                }

                simulatedShapes = DuplicateShapes(selectedShapes);
                if (PositionsLabMain.AlignReference == PositionsLabMain.AlignReferenceObject.PowerpointDefaults)
                {
                    positionsAction.Invoke(selectedShapes, dimension);
                }
                else
                {
                    var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                    var initialPositions = SaveOriginalPositions(simulatedPPShapes);

                    positionsAction.Invoke(simulatedShapes, dimension);

                    SyncShapes(selectedShapes, simulatedShapes, initialPositions);
                }
                if (isPreview)
                {
                    _previewIsExecuted = true;
                }
            }
            catch (Exception ex)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(ex.Message, ex);
                }
            }
            finally
            {
                if (simulatedShapes != null)
                {
                    simulatedShapes.Delete();
                    GC.Collect();
                }
            }
        }

        // Align center
        public void ExecutePositionsAction(Action<PowerPoint.ShapeRange, float, float> positionsAction, float dimension1, float dimension2, bool isPreview)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(ErrorMessageNoSelection);
                }
                return;
            }

            PowerPoint.ShapeRange simulatedShapes = null;

            try
            {
                var selectedShapes = this.GetCurrentSelection().ShapeRange;

                if (isPreview)
                {
                    SaveSelectedShapePositions(selectedShapes, allShapePos);
                }
                else
                {
                    UndoPreview();
                    _previewCallBack = null;
                    this.StartNewUndoEntry();
                }

                simulatedShapes = DuplicateShapes(selectedShapes);
                if (PositionsLabMain.AlignReference == PositionsLabMain.AlignReferenceObject.PowerpointDefaults)
                {
                    positionsAction.Invoke(selectedShapes, dimension1, dimension2);
                }
                else
                {
                    var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                    var initialPositions = SaveOriginalPositions(simulatedPPShapes);

                    positionsAction.Invoke(simulatedShapes, dimension1, dimension2);

                    SyncShapes(selectedShapes, simulatedShapes, initialPositions);
                }

                if (isPreview)
                {
                    _previewIsExecuted = true;
                }
            }
            catch (Exception ex)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(ex.Message, ex);
                }
            }
            finally
            {
                if (simulatedShapes != null)
                {
                    simulatedShapes.Delete();
                    GC.Collect();
                }
            }
        }

        public void ExecutePositionsAction(Action<List<PPShape>> positionsAction, bool isPreview)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(ErrorMessageNoSelection);
                }
                return;
            }

            PowerPoint.ShapeRange simulatedShapes = null;

            try
            {
                var selectedShapes = this.GetCurrentSelection().ShapeRange;

                if (isPreview)
                {
                    SaveSelectedShapePositions(selectedShapes, allShapePos);
                }
                else
                {
                    UndoPreview();
                    _previewCallBack = null;
                    this.StartNewUndoEntry();
                }

                simulatedShapes = DuplicateShapes(selectedShapes);
                var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                var initialPositions = SaveOriginalPositions(simulatedPPShapes);

                positionsAction.Invoke(simulatedPPShapes);

                SyncShapes(selectedShapes, simulatedPPShapes, initialPositions);

                if (isPreview)
                {
                    _previewIsExecuted = true;
                }
            }
            catch (Exception ex)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(ex.Message, ex);
                }
            }
            finally
            {
                if (simulatedShapes != null)
                {
                    simulatedShapes.Delete();
                    GC.Collect();
                }
            }
        }

        public void ExecutePositionsAction(Action<List<PPShape>, bool> positionsAction, bool booleanVal, bool isPreview)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(ErrorMessageNoSelection);
                }
                return;
            }

            PowerPoint.ShapeRange simulatedShapes = null;

            try
            {
                var selectedShapes = this.GetCurrentSelection().ShapeRange;

                if (isPreview)
                {
                    SaveSelectedShapePositions(selectedShapes, allShapePos);
                }
                else
                {
                    UndoPreview();
                    _previewCallBack = null;
                    this.StartNewUndoEntry();
                }

                simulatedShapes = DuplicateShapes(selectedShapes);
                var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                var initialPositions = SaveOriginalPositions(simulatedPPShapes);

                positionsAction.Invoke(simulatedPPShapes, booleanVal);

                SyncShapes(selectedShapes, simulatedPPShapes, initialPositions);

                if (isPreview)
                {
                    _previewIsExecuted = true;
                }
            }
            catch (Exception ex)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(ex.Message, ex);
                }
            }
            finally
            {
                if (simulatedShapes != null)
                {
                    simulatedShapes.Delete();
                    GC.Collect();
                }
            }
        }

        public void ExecutePositionsAction(Action<List<PPShape>, float> positionsAction, float dimension, bool isPreview)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(ErrorMessageNoSelection);
                }
                return;
            }

            PowerPoint.ShapeRange simulatedShapes = null;

            try
            {
                var selectedShapes = this.GetCurrentSelection().ShapeRange;

                if (isPreview)
                {
                    SaveSelectedShapePositions(selectedShapes, allShapePos);
                }
                else
                {
                    UndoPreview();
                    _previewCallBack = null;
                    this.StartNewUndoEntry();
                }

                simulatedShapes = DuplicateShapes(selectedShapes);
                var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                var initialPositions = SaveOriginalPositions(simulatedPPShapes);

                positionsAction.Invoke(simulatedPPShapes, dimension);

                SyncShapes(selectedShapes, simulatedPPShapes, initialPositions);

                if (isPreview)
                {
                    _previewIsExecuted = true;
                }
            }
            catch (Exception ex)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(ex.Message, ex);
                }
            }
            finally
            {
                if (simulatedShapes != null)
                {
                    simulatedShapes.Delete();
                    GC.Collect();
                }
            }
        }

        public void ExecutePositionsAction(Action<List<PPShape>, float, float> positionsAction, float dimension1, float dimension2, bool isPreview)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(ErrorMessageNoSelection);
                }
                return;
            }

            PowerPoint.ShapeRange simulatedShapes = null;

            try
            {
                var selectedShapes = this.GetCurrentSelection().ShapeRange;

                if (isPreview)
                {
                    SaveSelectedShapePositions(selectedShapes, allShapePos);
                }
                else
                {
                    UndoPreview();
                    _previewCallBack = null;
                    this.StartNewUndoEntry();
                }

                simulatedShapes = DuplicateShapes(selectedShapes);
                var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                var initialPositions = SaveOriginalPositions(simulatedPPShapes);

                positionsAction.Invoke(simulatedPPShapes, dimension1, dimension2);

                SyncShapes(selectedShapes, simulatedPPShapes, initialPositions);

                if (isPreview)
                {
                    _previewIsExecuted = true;
                }
            }
            catch (Exception ex)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(ex.Message, ex);
                }
            }
            finally
            {
                if (simulatedShapes != null)
                {
                    simulatedShapes.Delete();
                    GC.Collect();
                }
            }
        }

        public void ExecutePositionsAction(Action<List<Shape>> positionsAction, bool isPreview)
        {
            if (this.GetCurrentSelection().Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(ErrorMessageNoSelection);
                }
                return;
            }

            try
            {
                var selectedShapes = this.GetCurrentSelection().ShapeRange;

                if (isPreview)
                {
                    SaveSelectedShapePositions(selectedShapes, allShapePos);
                }
                else
                {
                    UndoPreview();
                    _previewCallBack = null;
                    this.StartNewUndoEntry();
                }

                positionsAction.Invoke(ConvertShapeRangeToShapeList(selectedShapes, 1));

                GC.Collect();

                if (isPreview)
                {
                    _previewIsExecuted = true;
                }
            }
            catch (Exception ex)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(ex.Message, ex);
                }
            }
        }

        private void PreviewHandler()
        {
            Focus();
            if (IsPreviewKeyPressed())
            {
                _previewCallBack.Invoke();
            }
        }

        private void UndoPreview(object sender, System.Windows.Input.MouseEventArgs e)
        {
            UndoPreview();
            _previewCallBack = null;
        }

        private void UndoPreview()
        {
            if (_previewIsExecuted)
            {
                var selectedShapes = this.GetCurrentSelection().ShapeRange;

                foreach (Shape s in selectedShapes)
                {
                    PositionShapeProperties properties;
                    var isPresent = allShapePos.TryGetValue(s.Id, out properties);

                    if (isPresent)
                    {
                        s.Left = properties.Position.X;
                        s.Top = properties.Position.Y;
                        s.Rotation = properties.Rotation;
                    }
                }

                _previewIsExecuted = false;
                GC.Collect();
            }
        }

        private bool IsPreviewKeyPressed()
        {
            if (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private bool IsChangeIconKeyPressed()
        {
            if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void SyncShapes(PowerPoint.ShapeRange selected, PowerPoint.ShapeRange simulatedShapes)
        {
            for (int i = 1; i <= selected.Count; i++)
            {
                var selectedShape = selected[i];
                var simulatedShape = simulatedShapes[i];

                selectedShape.IncrementLeft(Graphics.GetCenterPoint(simulatedShape).X - Graphics.GetCenterPoint(selectedShape).X);
                selectedShape.IncrementTop(Graphics.GetCenterPoint(simulatedShape).Y - Graphics.GetCenterPoint(selectedShape).Y);
                selectedShape.Rotation = simulatedShape.Rotation;
            }
        }

        private void SyncShapes(PowerPoint.ShapeRange selected, PowerPoint.ShapeRange simulatedShapes, float[,] originalPositions)
        {
            for (int i = 1; i <= selected.Count; i++)
            {
                var selectedShape = selected[i];
                var simulatedShape = simulatedShapes[i];

                selectedShape.IncrementLeft(Graphics.GetCenterPoint(simulatedShape).X - originalPositions[i - 1, Left]);
                selectedShape.IncrementTop(Graphics.GetCenterPoint(simulatedShape).Y - originalPositions[i - 1, Top]);
            }
        }

        private void SyncShapes(PowerPoint.ShapeRange selected, List<PPShape> simulatedShapes, float[,] originalPositions)
        {
            for (int i = 1; i <= selected.Count; i++)
            {
                var selectedShape = selected[i];
                var simulatedShape = simulatedShapes[i - 1];

                selectedShape.IncrementLeft(simulatedShape.VisualCenter.X - originalPositions[i - 1, Left]);
                selectedShape.IncrementTop(simulatedShape.VisualCenter.Y - originalPositions[i - 1, Top]);
            }
        }

        private PowerPoint.ShapeRange DuplicateShapes(PowerPoint.ShapeRange range)
        {
            int totalShapes = this.GetCurrentSlide().Shapes.Count;
            int[] duplicatedShapeIndices = new int[range.Count];

            for (int i = 1; i <= range.Count; i++)
            {
                var shape = range[i];
                var duplicated = shape.Duplicate()[1];
                duplicated.Name = shape.Id + "";
                duplicated.Left = shape.Left;
                duplicated.Top = shape.Top;
                duplicatedShapeIndices[i - 1] = totalShapes + i;
            }

            return this.GetCurrentSlide().Shapes.Range(duplicatedShapeIndices);
        }

        private float[,] SaveOriginalPositions(List<PPShape> shapes)
        {
            var initialPositions = new float[shapes.Count, 2];
            for (var i = 0; i < shapes.Count; i++)
            {
                var s = shapes[i];
                var pt = s.VisualCenter;
                initialPositions[i, Left] = pt.X;
                initialPositions[i, Top] = pt.Y;
            }

            return initialPositions;
        }

        private void SaveSelectedShapePositions(PowerPoint.ShapeRange shapes, Dictionary<int, PositionShapeProperties> dictionary)
        {
            dictionary.Clear();
            foreach (Shape s in shapes)
            {
                dictionary.Add(s.Id, new PositionShapeProperties(new System.Drawing.PointF(s.Left, s.Top), s.Rotation, s.HorizontalFlip, s.VerticalFlip));
            }
        }
        #endregion
    }
}
