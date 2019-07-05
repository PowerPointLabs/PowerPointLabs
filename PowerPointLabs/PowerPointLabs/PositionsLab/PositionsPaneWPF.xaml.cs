using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Threading;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.PositionsLab.Views;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

using PPExtraEventHelper;

using Office = Microsoft.Office.Core;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.PositionsLab
{
    [SuppressMessage("Microsoft.StyleCop.CSharp.OrderingRules", "SA1202:ElementsMustBeOrderedByAccess", Justification = "To refactor to partials")]
    /// <summary>
    /// Interaction logic for PositionsPaneWPF.xaml
    /// </summary>
    public partial class PositionsPaneWpf
    {
        private PositionsDistributeGridDialog _positionsDistributeGridDialog;

        #pragma warning disable 0618
        private LMouseUpListener _leftMouseUpListener;
        private LMouseDownListener _leftMouseDownListener;
        private DispatcherTimer _dispatcherTimer;

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
        private static ShapeRange _selectedRange;

        //Variables for key binding
        private const int CtrlProportion = 5;

        public PositionsPaneWpf()
        {
            PositionsLabMain.InitPositionsLab();
            InitializeComponent();
            InitializeHotKeys();
            _dispatcherTimer = new DispatcherTimer(DispatcherPriority.Background, Dispatcher);
            _dispatcherTimer.Interval = TimeSpan.FromMilliseconds(10);
            Focusable = true;

        }

        ~PositionsPaneWpf()
        {
            ClearAllEventHandlers();
        }

        private void InitializeHotKeys()
        {
            Dictionary<ToggleButton, Action<bool, bool>> buttonActionMapping = new Dictionary<ToggleButton, Action<bool, bool>>();
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
                Microsoft.Office.Tools.CustomTaskPane positionsPane = this.GetTaskPane(typeof(PositionsPane));
                if (positionsPane == null || !positionsPane.Visible)
                {
                    return false;
                }

                foreach (KeyValuePair<ToggleButton, Action<bool, bool>> mapping in buttonActionMapping)
                {
                    ToggleButton button = mapping.Key;
                    Action<bool, bool> action = mapping.Value;

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
            System.Drawing.PointF origin = _refPoint.GetCenterPoint();
            float angle = (isClockwise) ? 1f : -1f;
            if (isLarge)
            {
                angle *= CtrlProportion;
            }

            foreach (Shape currentShape in _shapesToBeRotated)
            {
                PositionsLabMain.Rotate(currentShape, origin, angle, PositionsLabSettings.ReorientShapeOrientation);
            }
        }

        #region Click Behaviour
        #region Align
        private void AlignLeftButton_Click(object sender, RoutedEventArgs e)
        {
            Action<ShapeRange> positionsAction = shapes => PositionsLabMain.AlignLeft(shapes);
            ExecutePositionsAction(positionsAction, false, true);
        }

        private void AlignRightButton_Click(object sender, RoutedEventArgs e)
        {
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            Action<ShapeRange, float> positionsAction = (shapes, width) => PositionsLabMain.AlignRight(shapes, width);
            ExecutePositionsAction(positionsAction, slideWidth, false);
        }

        private void AlignTopButton_Click(object sender, RoutedEventArgs e)
        {
            Action<ShapeRange> positionsAction = shapes => PositionsLabMain.AlignTop(shapes);
            ExecutePositionsAction(positionsAction, false, true);
        }

        private void AlignBottomButton_Click(object sender, RoutedEventArgs e)
        {
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<ShapeRange, float> positionsAction = (shapes, height) => PositionsLabMain.AlignBottom(shapes, height);
            ExecutePositionsAction(positionsAction, slideHeight, false);
        }

        private void AlignHorizontalCenterButton_Click(object sender, RoutedEventArgs e)
        {
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<ShapeRange, float> positionsAction = (shapes, height) => PositionsLabMain.AlignHorizontalCenter(shapes, height);
            ExecutePositionsAction(positionsAction, slideHeight, false);
        }

        private void AlignVerticalCenterButton_Click(object sender, RoutedEventArgs e)
        {
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            Action<ShapeRange, float> positionsAction = (shapes, width) => PositionsLabMain.AlignVerticalCenter(shapes, width);
            ExecutePositionsAction(positionsAction, slideWidth, false);
        }

        private void AlignCenterButton_Click(object sender, RoutedEventArgs e)
        {
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            Action<ShapeRange, float, float> positionsAction = (shapes, height, width) => PositionsLabMain.AlignCenter(shapes, height, width);
            ExecutePositionsAction(positionsAction, slideHeight, slideWidth, false);
        }

        private void AlignRadialButton_Click(object sender, RoutedEventArgs e)
        {
            Action<ShapeRange> positionsAction = (shapes) => PositionsLabMain.AlignRadial(shapes);
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
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            Action<List<PPShape>, float> positionsAction = (shapes, width) => PositionsLabMain.DistributeHorizontal(shapes, width);
            ExecutePositionsAction(positionsAction, slideWidth, false);
        }

        private void DistributeVerticalButton_Click(object sender, RoutedEventArgs e)
        {
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<List<PPShape>, float> positionsAction = (shapes, height) => PositionsLabMain.DistributeVertical(shapes, height);
            ExecutePositionsAction(positionsAction, slideHeight, false);
        }

        private void DistributeCenterButton_Click(object sender, RoutedEventArgs e)
        {
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<List<PPShape>, float, float> positionsAction = (shapes, width, height) => PositionsLabMain.DistributeCenter(shapes, width, height);
            ExecutePositionsAction(positionsAction, slideWidth, slideHeight, false);
        }
        
        private void DistributeGridButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.GetCurrentSelection().Type != PpSelectionType.ppSelectionShapes)
            {
                ShowErrorMessageBox(PositionsLabText.ErrorNoSelection);
                return;
            }

            try
            {
                ShapeRange selectedShapes = this.GetCurrentSelection().ShapeRange;
                int numShapesSelected = selectedShapes.Count;
                int colLength = (int)Math.Ceiling(Math.Sqrt(numShapesSelected));
                int rowLength = (int)Math.Ceiling((double)numShapesSelected / colLength);

                _positionsDistributeGridDialog = new PositionsDistributeGridDialog(numShapesSelected, rowLength, colLength,
                                                                                    PositionsLabSettings.DistributeGridAlignment,
                                                                                    PositionsLabSettings.GridMarginTop,
                                                                                    PositionsLabSettings.GridMarginBottom,
                                                                                    PositionsLabSettings.GridMarginLeft,
                                                                                    PositionsLabSettings.GridMarginRight);
                _positionsDistributeGridDialog.DialogConfirmedHandler += OnDistributeGridDialogConfirmed;
                _positionsDistributeGridDialog.ShowThematicDialog();
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }  
        }

        private void OnDistributeGridDialogConfirmed(int rowLength, int colLength,
                                        PositionsLabSettings.GridAlignment gridAlignment,
                                        float gridMarginTop, float gridMarginBottom,
                                        float gridMarginLeft, float gridMarginRight)
        {
            PositionsLabSettings.DistributeGridAlignment = gridAlignment;
            PositionsLabSettings.GridMarginTop = gridMarginTop;
            PositionsLabSettings.GridMarginBottom = gridMarginBottom;
            PositionsLabSettings.GridMarginLeft = gridMarginLeft;
            PositionsLabSettings.GridMarginRight = gridMarginRight;

            ExecuteDistributeGrid(rowLength, colLength);
        }

        private void DistributeRadialButton_Click(object sender, RoutedEventArgs e)
        {
            Action<ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
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
            bool noShapesSelected = this.GetCurrentSelection().Type != PpSelectionType.ppSelectionShapes;
            ToggleButton button = (ToggleButton)sender;

            if (noShapesSelected)
            {
                button.IsChecked = false;
                ShowErrorMessageBox(PositionsLabText.ErrorFewerThanTwoSelection);
                return;
            }

            ShapeRange selectedShapes = this.GetCurrentSelection().ShapeRange;

            if (selectedShapes.Count <= 1)
            {
                button.IsChecked = false;
                ShowErrorMessageBox(PositionsLabText.ErrorFewerThanTwoSelection);
                return;
            }

            ClearAllEventHandlers();

            Models.PowerPointSlide currentSlide = this.GetCurrentSlide();

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
            System.Drawing.Point p = System.Windows.Forms.Control.MousePosition;

            float prevAngle = (float)PositionsLabMain.AngleBetweenTwoPoints(ConvertSlidePointToScreenPoint(_refPoint.GetCenterPoint()), _prevMousePos);
            float angle = (float)PositionsLabMain.AngleBetweenTwoPoints(ConvertSlidePointToScreenPoint(_refPoint.GetCenterPoint()), p) - prevAngle;

            System.Drawing.PointF origin = _refPoint.GetCenterPoint();

            foreach (Shape currentShape in _shapesToBeRotated)
            {
                PositionsLabMain.Rotate(currentShape, origin, angle, PositionsLabSettings.ReorientShapeOrientation);
            }

            _prevMousePos = p;
        }

        void _leftMouseUpListener_Rotation(object sender, SysMouseEventInfo e)
        {
            try
            {
                _dispatcherTimer.Stop();
                _selectedRange.Select();
            }
            catch (Exception ex)
            {
                rotationButton.IsChecked = false;
                duplicateRotationButton.IsChecked = false;
                Logger.LogException(ex, "Rotation");
            }
        }

        void _leftMouseDownListener_Rotation(object sender, SysMouseEventInfo e)
        {
            ToggleButton button = ((bool)rotationButton.IsChecked) ? rotationButton :
                             ((bool)duplicateRotationButton.IsChecked) ? duplicateRotationButton
                                                                       : null;
            try
            {
                if (button == null || button.IsMouseOver)
                {
                    DisableRotationMode();
                    return;
                }

                System.Drawing.Point p = System.Windows.Forms.Control.MousePosition;
                Shape selectedShape = GetShapeDirectlyBelowMousePos(_allShapesInSlide, p);

                if (selectedShape == null)
                {
                    DisableRotationMode();
                    button.IsChecked = false;
                    return;
                }

                bool isShapeToBeRotated = _shapesToBeRotated.Contains(selectedShape);
                bool isRefPoint = _refPoint.Id == selectedShape.Id;

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
                    foreach (Shape currentShape in _shapesToBeRotated)
                    {
                        Shape duplicatedShape = currentShape.Duplicate()[1];
                        duplicatedShape.Left -= 12;
                        duplicatedShape.Top -= 12;
                        ShapeUtil.MoveZToJustBehind(duplicatedShape, currentShape);
                    }
                }

                _prevMousePos = p;
                _dispatcherTimer.Start();
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "Rotation");
                DisableRotationMode();
                button.IsChecked = false;
            }
        }

        private void LockAxisButton_Click(object sender, RoutedEventArgs e)
        {
            bool noShapesSelected = this.GetCurrentSelection().Type != PpSelectionType.ppSelectionShapes;

            if (noShapesSelected)
            {
                lockAxisButton.IsChecked = false;
                ShowErrorMessageBox(PositionsLabText.ErrorNoSelection);
                return;
            }

            ShapeRange selectedShapes = this.GetCurrentSelection().ShapeRange;

            ClearAllEventHandlers();

            Models.PowerPointSlide currentSlide = this.GetCurrentSlide();

            _shapesToBeMoved = ConvertShapeRangeToShapeList(selectedShapes, 1);
            _allShapesInSlide = ConvertShapesToShapeList(currentSlide.Shapes);
            _selectedRange = selectedShapes;

            StartLockAxisMode();
        }

        private void LockAxisHandler(object sender, EventArgs e)
        {
            //Remove dragging control of user
            this.GetCurrentSelection().Unselect();

            System.Drawing.Point currentMousePos = System.Windows.Forms.Control.MousePosition;

            float diffX = currentMousePos.X - _initialMousePos.X;
            float diffY = currentMousePos.Y - _initialMousePos.Y;

            for (int i = 0; i < _shapesToBeMoved.Count; i++)
            {
                Shape s = _shapesToBeMoved[i];
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

                System.Drawing.Point p = System.Windows.Forms.Control.MousePosition;
                Models.PowerPointSlide currentSlide = this.GetCurrentSlide();
                Shape selectedShape = GetShapeDirectlyBelowMousePos(_allShapesInSlide, p);

                if (selectedShape == null)
                {
                    DisableLockAxisMode();
                    lockAxisButton.IsChecked = false;
                    return;
                }

                bool isShapeToBeMoved = _shapesToBeMoved.Contains(selectedShape);

                if (!isShapeToBeMoved)
                {
                    DisableLockAxisMode();
                    lockAxisButton.IsChecked = false;
                    return;
                }

                this.StartNewUndoEntry();

                _initialPos = new float[_shapesToBeMoved.Count, 2];
                for (int i = 0; i < _shapesToBeMoved.Count; i++)
                {
                    Shape s = _shapesToBeMoved[i];
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
            Action<ShapeRange> positionsAction = shapes => PositionsLabMain.AlignLeft(shapes);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, true, true);
            };
            PreviewHandler();
        }

        private void AlignRightButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            Action<ShapeRange, float> positionsAction = (shapes, width) => PositionsLabMain.AlignRight(shapes, width);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, slideWidth, true);
            };
            PreviewHandler();
        }

        private void AlignTopButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            Action<ShapeRange> positionsAction = shapes => PositionsLabMain.AlignTop(shapes);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, true, true);
            };
            PreviewHandler();
        }

        private void AlignBottomButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<ShapeRange, float> positionsAction = (shapes, height) => PositionsLabMain.AlignBottom(shapes, height);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, slideHeight, true);
            };
            PreviewHandler();
        }

        private void AlignHorizontalCenterButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<ShapeRange, float> positionsAction = (shapes, height) => PositionsLabMain.AlignHorizontalCenter(shapes, height);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, slideHeight, true);
            };
            PreviewHandler();
        }

        private void AlignVerticalCenterButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            Action<ShapeRange, float> positionsAction = (shapes, width) => PositionsLabMain.AlignVerticalCenter(shapes, width);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, slideWidth, true);
            };
            PreviewHandler();
        }

        private void AlignCenterButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            Action<ShapeRange, float, float> positionsAction = (shapes, height, width) => PositionsLabMain.AlignCenter(shapes, height, width);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, slideHeight, slideWidth, true);
            };
            PreviewHandler();
        }

        private void AlignRadialButton_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<ShapeRange> positionsAction = (shapes) => PositionsLabMain.AlignRadial(shapes);
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
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            Action<List<PPShape>, float> positionsAction = (shapes, width) => PositionsLabMain.DistributeHorizontal(shapes, width);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, slideWidth, true);
            };
            PreviewHandler();
        }

        private void DistributeVerticalButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<List<PPShape>, float> positionsAction = (shapes, height) => PositionsLabMain.DistributeVertical(shapes, height);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, slideHeight, true);
            };
            PreviewHandler();
        }

        private void DistributeCenterButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<List<PPShape>, float, float> positionsAction = (shapes, width, height) => PositionsLabMain.DistributeCenter(shapes, width, height);
            _previewCallBack = delegate
            {
                ExecutePositionsAction(positionsAction, slideWidth, slideHeight, true);
            };
            PreviewHandler();
        }

        private void DistributeRadialButton_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
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
        private Shape AddReferencePoint(Shapes shapes, float left, float top)
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
            float epsilon = 0.00001f;

            System.Drawing.PointF centerPoint = ConvertSlidePointToScreenPoint(shape.GetCenterPoint());
            System.Drawing.PointF rotatedMousePos = CommonUtil.RotatePoint(p, centerPoint, -shape.Rotation);

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

        private List<PPShape> ConvertShapeRangeToPPShapeList (ShapeRange range, int index)
        {
            List<PPShape> shapes = new List<PPShape>();

            for (int i = index; i <= range.Count; i++)
            {
                Shape s = range[i];
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

        private List<Shape> ConvertShapeRangeToShapeList(ShapeRange range, int index)
        {
            List<Shape> shapes = new List<Shape>();

            for (int i = index; i <= range.Count; i++)
            {
                shapes.Add(range[i]);
            }

            return shapes;
        }

        private List<Shape> ConvertShapesToShapeList(Shapes shapes)
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

        #region Settings Dialog
        private void AlignSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabSettings.ShowAlignSettingsDialog();
        }

        private void DistributeSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabSettings.ShowDistributeSettingsDialog();
        }

        private void ReorderSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabSettings.ShowReorderSettingsDialog();
        }

        private void ReorientSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabSettings.ShowReorientSettingsDialog();
        }
        #endregion

        public void ClearAllEventHandlers()
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
            _dispatcherTimer = new DispatcherTimer(DispatcherPriority.Background, Dispatcher);
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
                MessageBox.Show(content, PositionsLabText.ErrorDialogTitle);
                return;
            }
            
            string errorMessage = GetErrorMessage(exception.Message);
            if (!string.Equals(errorMessage, PositionsLabText.ErrorUndefined, StringComparison.Ordinal))
            {
                MessageBox.Show(content, PositionsLabText.ErrorDialogTitle);
            }
            else
            {
                ErrorDialogBox.ShowDialog(PositionsLabText.ErrorDialogTitle, content, exception);
            }
        }

        private string GetErrorMessage(string errorMsg)
        {
            switch (errorMsg)
            {
                case PositionsLabText.ErrorNoSelection:
                    return PositionsLabText.ErrorNoSelection;
                case PositionsLabText.ErrorFewerThanTwoSelection:
                    return PositionsLabText.ErrorFewerThanTwoSelection;
                case PositionsLabText.ErrorFewerThanThreeSelection:
                    return PositionsLabText.ErrorFewerThanThreeSelection;
                case PositionsLabText.ErrorFewerThanFourSelection:
                    return PositionsLabText.ErrorFewerThanFourSelection;
                case PositionsLabText.ErrorFunctionNotSupportedForWithinShapes:
                    return PositionsLabText.ErrorFunctionNotSupportedForWithinShapes;
                case PositionsLabText.ErrorFunctionNotSupportedForSlide:
                    return PositionsLabText.ErrorFunctionNotSupportedForSlide;
                case PositionsLabText.ErrorFunctionNotSupportedForOverlapRefShapeCenter:
                    return PositionsLabText.ErrorFunctionNotSupportedForOverlapRefShapeCenter;
                default:
                    return PositionsLabText.ErrorUndefined;
            }
        }

        private void IgnoreExceptionThrown() { }

        #endregion

        #region Helper

        // returns true if selection is valid
        public bool HandleInvalidSelection(bool isPreview, Selection selection)
        {
            if (!ShapeUtil.IsSelectionShape(selection))
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(PositionsLabText.ErrorNoSelection);
                }
                return false;
            }
            return true;
        }
        // Align left and top, Align/Distribute radial
        public void ExecutePositionsAction(Action<ShapeRange> positionsAction, bool isPreview, bool isConvertPPShape)
        {
            Selection selection = this.GetCurrentSelection();
            if (!HandleInvalidSelection(isPreview, selection))
            {
                // invalid selection!
                return;
            }

            ShapeRange simulatedShapes = null;

            try
            {
                ShapeRange selectedShapes = selection.ShapeRange;

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

                // selectedShapes.Duplicate() may return a list with reversed sequence  
                simulatedShapes = DuplicateShapes(selectedShapes); 

                if (PositionsLabSettings.AlignReference == PositionsLabSettings.AlignReferenceObject.PowerpointDefaults)
                {
                    positionsAction.Invoke(selectedShapes);
                }
                else if (isConvertPPShape)
                {
                    List<PPShape> simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                    float[,] initialPositions = SaveOriginalPositions(simulatedPPShapes);

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
                    try
                    {
                        simulatedShapes.Delete();
                        GC.Collect();
                    }
                    // Catch corrupted shapes
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Remove all simulated shapes manually
                        for (int i = 0; i < simulatedShapes.Count; i++)
                        {
                            // This method to remove duplicated shapes might fail for non-corrupted shapes when mixing good/bad shapes
                            if (this.GetCurrentSlide().Shapes[1].Name.Contains("_Copy"))
                            {
                                this.GetCurrentSlide().Shapes[1].Delete();
                            }
                            else
                            {
                                break;
                            }
                        }
                        // Remove any outlier extra shapes not deleted previously
                        // Only triggered for cases where Distribute is called for cases consisting of both
                        // corrupted and non-corrupted shapes
                        try
                        {
                            this.GetCurrentSelection().Delete();
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            // Exception will trigger whenever Distribute is applied to cases where all shapes are 
                            // either corrupted or non-corrupted, which is already handled before this try-catch block
                        }
                        // Ask user to undo the operation to remove any excess duplicates
                        MessageBox.Show(PositionsLabText.ErrorCorruptedSelection, PositionsLabText.ErrorCorruptedShapesTitle);
                    }
                }
            }
        }

        // Align right, bottom, vertical center, horizontal center
        public void ExecutePositionsAction(Action<ShapeRange, float> positionsAction, float dimension, bool isPreview)
        {
            Selection selection = this.GetCurrentSelection();
            if (!HandleInvalidSelection(isPreview, selection))
            {
                // invalid selection!
                return;
            }

            ShapeRange simulatedShapes = null;

            try
            {
                ShapeRange selectedShapes = selection.ShapeRange;

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
                if (PositionsLabSettings.AlignReference == PositionsLabSettings.AlignReferenceObject.PowerpointDefaults)
                {
                    positionsAction.Invoke(selectedShapes, dimension);
                }
                else
                {
                    List<PPShape> simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                    float[,] initialPositions = SaveOriginalPositions(simulatedPPShapes);

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
        public void ExecutePositionsAction(Action<ShapeRange, float, float> positionsAction, float dimension1, float dimension2, bool isPreview)
        {
            Selection selection = this.GetCurrentSelection();
            if (!HandleInvalidSelection(isPreview, selection))
            {
                // invalid selection!
                return;
            }

            ShapeRange simulatedShapes = null;

            try
            {
                ShapeRange selectedShapes = selection.ShapeRange;

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
                simulatedShapes.Select();
                if (PositionsLabSettings.AlignReference == PositionsLabSettings.AlignReferenceObject.PowerpointDefaults)
                {
                    positionsAction.Invoke(selectedShapes, dimension1, dimension2);
                }
                else
                {
                    List<PPShape> simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                    float[,] initialPositions = SaveOriginalPositions(simulatedPPShapes);

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
        // Adjoin operations
        public void ExecutePositionsAction(Action<List<PPShape>> positionsAction, bool isPreview)
        {
            // Need to run the action 2 times because of the nature of PowerPoint default operations
            // This has been determined via manual testing
            for (int i = 0; i < 2; i++)
            {
                Selection selection = this.GetCurrentSelection();
                if (!HandleInvalidSelection(isPreview, selection))
                {
                    // invalid selection!
                    return;
                }

                ShapeRange simulatedShapes = null;

                try
                {
                    ShapeRange selectedShapes = selection.ShapeRange;
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
                    simulatedShapes = selectedShapes.Duplicate();
                    List<PPShape> simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                    float[,] initialPositions = SaveOriginalPositions(simulatedPPShapes);

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
                        break;
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
        }

        public void ExecutePositionsAction(Action<List<PPShape>, bool> positionsAction, bool booleanVal, bool isPreview)
        {
            Selection selection = this.GetCurrentSelection();
            if (!HandleInvalidSelection(isPreview, selection))
            {
                // invalid selection!
                return;
            }

            ShapeRange simulatedShapes = null;
            try
            {
                ShapeRange selectedShapes = selection.ShapeRange;

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

                // set the zOrder
                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    PPShape simulatedPPShape = new PPShape(simulatedShapes[i], false);
                    ShapeUtil.SwapZOrder(simulatedPPShape._shape, selectedShapes[i]);
                }

                List<PPShape> simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                float[,] initialPositions = SaveOriginalPositions(simulatedPPShapes);

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

        // Distribute horizontal and vertical
        public void ExecutePositionsAction(Action<List<PPShape>, float> positionsAction, float dimension, bool isPreview)
        {
            // Need to run the action 2 times because of the nature of PowerPoint default operations
            // This has been determined via manual testing
            for (int numOfRuns = 0; numOfRuns < 2; numOfRuns++)
            {
                Selection selection = this.GetCurrentSelection();
                if (!HandleInvalidSelection(isPreview, selection))
                {
                    // invalid selection!
                    return;
                }

                ShapeRange simulatedShapes = null;

                try
                {
                    ShapeRange selectedShapes = selection.ShapeRange;

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

                    simulatedShapes = selectedShapes.Duplicate();
                    List<PPShape> simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                    float[,] initialPositions = SaveOriginalPositions(simulatedPPShapes);

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
                        break;
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
        }

        // Distribute center
        public void ExecutePositionsAction(Action<List<PPShape>, float, float> positionsAction, float dimension1, float dimension2, bool isPreview)
        {
            // Need to run the action 2 times because of the nature of PowerPoint default operations
            // This has been determined via manual testing
            for (int numOfRuns = 0; numOfRuns < 2; numOfRuns++)
            {
                if (this.GetCurrentSelection().Type != PpSelectionType.ppSelectionShapes)
                {
                    if (!isPreview)
                    {
                        ShowErrorMessageBox(PositionsLabText.ErrorNoSelection);
                    }
                    return;
                }

                ShapeRange simulatedShapes = null;

                try
                {
                    ShapeRange selectedShapes = this.GetCurrentSelection().ShapeRange;

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

                    simulatedShapes = selectedShapes.Duplicate();
                    List<PPShape> simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                    float[,] initialPositions = SaveOriginalPositions(simulatedPPShapes);

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
                        break;
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
        }

        public void ExecutePositionsAction(Action<List<Shape>> positionsAction, bool isPreview)
        {
            if (this.GetCurrentSelection().Type != PpSelectionType.ppSelectionShapes)
            {
                if (!isPreview)
                {
                    ShowErrorMessageBox(PositionsLabText.ErrorNoSelection);
                }
                return;
            }

            try
            {
                ShapeRange selectedShapes = this.GetCurrentSelection().ShapeRange;

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

        // Distribute grid
        private void ExecuteDistributeGrid(int rowLength, int colLength)
        {
            ShapeRange simulatedShapes = null;
            try
            {
                this.StartNewUndoEntry();
                ShapeRange selectedShapes = this.GetCurrentSelection().ShapeRange;
                simulatedShapes = DuplicateShapes(selectedShapes);
                List<PPShape> simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);

                PositionsLabMain.DistributeGrid(simulatedPPShapes, rowLength, colLength);

                SyncShapes(selectedShapes, simulatedShapes);
            }
            catch (Exception ex)
            {
                ShowErrorMessageBox(ex.Message, ex);
            }
            finally
            {
                if (simulatedShapes != null)
                {
                    try
                    {
                        simulatedShapes.Delete();
                        GC.Collect();
                    }
                    // Catch corrupted shapes
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Remove all simulated shapes manually
                        for (int i = 0; i < simulatedShapes.Count; i++)
                        {
                            // This method to remove duplicated shapes might fail for non-corrupted shapes when mixing good/bad shapes
                            if (this.GetCurrentSlide().Shapes[1].Name.Contains("_Copy"))
                            {
                                this.GetCurrentSlide().Shapes[1].Delete();
                            }
                            else
                            {
                                break;
                            }
                        }
                        // Remove any outlier extra shapes not deleted previously
                        // Only triggered for cases where Distribute is called for cases consisting of both
                        // corrupted and non-corrupted shapes
                        try
                        {
                            this.GetCurrentSelection().Delete();
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            // Exception will trigger whenever Distribute is applied to cases where all shapes are 
                            // either corrupted or non-corrupted, which is already handled before this try-catch block
                        }
                        // Ask user to undo the operation to remove any excess duplicates
                        MessageBox.Show(PositionsLabText.ErrorCorruptedSelection, PositionsLabText.ErrorCorruptedShapesTitle);
                    }
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

        private void UndoPreview(object sender, MouseEventArgs e)
        {
            UndoPreview();
            _previewCallBack = null;
        }

        private void UndoPreview()
        {
            if (_previewIsExecuted)
            {
                ShapeRange selectedShapes = this.GetCurrentSelection().ShapeRange;

                foreach (Shape s in selectedShapes)
                {
                    PositionShapeProperties properties;
                    bool isPresent = allShapePos.TryGetValue(s.Id, out properties);

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

        private void SyncShapes(ShapeRange selected, ShapeRange simulatedShapes)
        {
            for (int i = 1; i <= selected.Count; i++)
            {
                Shape selectedShape = selected[i];
                Shape simulatedShape = simulatedShapes[i];

                selectedShape.IncrementLeft(simulatedShape.GetCenterPoint().X - selectedShape.GetCenterPoint().X);
                selectedShape.IncrementTop(simulatedShape.GetCenterPoint().Y - selectedShape.GetCenterPoint().Y);
                selectedShape.Rotation = simulatedShape.Rotation;
            }
        }

        private void SyncShapes(ShapeRange selected, ShapeRange simulatedShapes, float[,] originalPositions)
        {
            for (int i = 1; i <= selected.Count; i++)
            {
                Shape selectedShape = selected[i];
                Shape simulatedShape = simulatedShapes[i];

                selectedShape.IncrementLeft(simulatedShape.GetCenterPoint().X - originalPositions[i - 1, Left]);
                selectedShape.IncrementTop(simulatedShape.GetCenterPoint().Y - originalPositions[i - 1, Top]);
            }
        }

        private void SyncShapes(ShapeRange selected, List<PPShape> simulatedShapes, float[,] originalPositions)
        {
            for (int i = 1; i <= selected.Count; i++)
            {
                Shape selectedShape = selected[i];
                PPShape simulatedShape = simulatedShapes[i - 1];

                selectedShape.IncrementLeft(simulatedShape.VisualCenter.X - originalPositions[i - 1, Left]);
                selectedShape.IncrementTop(simulatedShape.VisualCenter.Y - originalPositions[i - 1, Top]);
                ShapeUtil.SwapZOrder(simulatedShape._shape, selectedShape);
            }
        }

        private ShapeRange DuplicateShapes(ShapeRange range)
        {
            String[] duplicatedShapeNames = new String[range.Count];
            for (int i = 0; i < range.Count; i++)
            {
                Shape shape = range[i + 1];
                Shape duplicated = shape.Duplicate()[1];

                // Add a number at end of name in case the name of shapes are same
                duplicated.Name = shape.Name + "_Copy_" + i.ToString();

                duplicated.Left = shape.Left;
                duplicated.Top = shape.Top;
                duplicatedShapeNames[i] = duplicated.Name;
            }
            return this.GetCurrentSlide().Shapes.Range(duplicatedShapeNames);
        }

        private float[,] SaveOriginalPositions(List<PPShape> shapes)
        {
            float[,] initialPositions = new float[shapes.Count, 2];
            for (int i = 0; i < shapes.Count; i++)
            {
                PPShape s = shapes[i];
                System.Drawing.PointF pt = s.VisualCenter;
                initialPositions[i, Left] = pt.X;
                initialPositions[i, Top] = pt.Y;
            }

            return initialPositions;
        }

        private void SaveSelectedShapePositions(ShapeRange shapes, Dictionary<int, PositionShapeProperties> dictionary)
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
