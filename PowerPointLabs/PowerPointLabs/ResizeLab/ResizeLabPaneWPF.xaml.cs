using System.Collections.Generic;
using System.Drawing;
using System.Windows;
using System.Windows.Input;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// Interaction logic for ResizePane.xaml
    /// </summary>
    public partial class ResizeLabPaneWPF : IResizeLabPane
    {
        private ResizeLabMain _resizeLab;
        public static bool IsAspectRatioLocked { get; set; }
        private const string UnlockText = "Unlocked";
        private const string LockText = "Locked";
        private const string UnlockAspectRatioToolTip = "Unlocks the aspect ratio of objects when performing resizing of objects";
        private const string LockAspectRatioToolTip = "Locks the aspect ratio of objects when performing resizing of objects";
        private readonly Bitmap _unlockedImage;
        private readonly Bitmap _lockedImage;
        private readonly Dictionary<string, ShapeProperties> _originalShapeProperties = new Dictionary<string, ShapeProperties>(); 

        public ResizeLabPaneWPF()
        {
            InitializeComponent();
            InitialiseLogicInstance();
            _unlockedImage = new Bitmap(Properties.Resources.ResizeUnlock);
            _lockedImage = new Bitmap(Properties.Resources.ResizeLock);
            UnlockAspectRatio();
        }

        internal void InitialiseLogicInstance()
        {
            if (_resizeLab == null)
            {
                _resizeLab = new ResizeLabMain(this);
            }
        }

        internal void InitialiseAspectRatio()
        {
            UnlockAspectRatio();
        }

        #region Event Handler: Strech and Shrink

        private void StretchLeftBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShape = GetSelectedShapes();
            var resizeAction = new SingleInputResizeAction(shapes => _resizeLab.StretchLeft);

            ExecuteResizeAction(selectedShape, resizeAction);
        }

        private void StretchRightBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            var resizeAction = new SingleInputResizeAction(shapes => _resizeLab.StretchRight);

            ExecuteResizeAction(selectedShapes, resizeAction);
        }

        private void StretchTopBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            var resizeAction = new SingleInputResizeAction(shapes => _resizeLab.StretchTop);

            ExecuteResizeAction(selectedShapes, resizeAction);
        }

        private void StretchBottomBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            var resizeAction = new SingleInputResizeAction(shapes => _resizeLab.StretchBottom);

            ExecuteResizeAction(selectedShapes, resizeAction);
        }

        #endregion

        #region Event Handler: Same Dimension

        private void SameWidthBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            var resizeAction = new SingleInputResizeAction(shapes => _resizeLab.ResizeToSameWidth);

            ExecuteResizeAction(selectedShapes, resizeAction);
        }

        private void SameHeightBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            var resizeAction = new SingleInputResizeAction(shapes => _resizeLab.ResizeToSameHeight);

            ExecuteResizeAction(selectedShapes, resizeAction);
        }

        private void SameSizeBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            var resizeAction = new SingleInputResizeAction(shapes => _resizeLab.ResizeToSameHeightAndWidth);

            ExecuteResizeAction(selectedShapes, resizeAction);
        }

        #endregion

        #region Event Handler: Fit
        private void FitWidthBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            var resizeAction = new MultiInputResizeAction((shapes, width, height, isAspectRatio) => _resizeLab.FitToWidth);

            ExecuteResizeAction(selectedShapes, slideWidth, slideHeight, resizeAction);
        }

        private void FitHeightBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            var resizeAction = new MultiInputResizeAction((shapes, width, height, isAspectRatio) => _resizeLab.FitToHeight);

            ExecuteResizeAction(selectedShapes, slideWidth, slideHeight, resizeAction);
        }

        private void FillBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            var resizeAction = new MultiInputResizeAction((shapes, width, height, isAspectRatio) => _resizeLab.FitToFill);

            ExecuteResizeAction(selectedShapes, slideWidth, slideHeight, resizeAction);
        }

        #endregion

        #region Event Handler: Aspect Ratio

        private void LockAspectRatioBtn_Click(object sender, RoutedEventArgs e)
        {
            if (IsAspectRatioLocked)
            {
                UnlockAspectRatio();
            }
            else
            {
                LockAspectRatio();
            }
        }

        private void RestoreAspectRatioBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            var slideWidth = this.GetCurrentPresentation().SlideWidth;

            if (selectedShapes != null)
            {
                _resizeLab.RestoreAspectRatio(selectedShapes, slideHeight, slideWidth);
            }
        }

        private void UnlockAspectRatio()
        {
            IsAspectRatioLocked = false;
            LockAspectRatioBtn.Text = UnlockText;
            LockAspectRatioBtn.ToolTip = LockAspectRatioToolTip;
            LockAspectRatioBtn.Image = Utils.Graphics.CreateBitmapSourceFromGdiBitmap(_unlockedImage);

            ModifySelectionAspectRatio();
        }

        private void LockAspectRatio()
        {
            IsAspectRatioLocked = true;
            LockAspectRatioBtn.Text = LockText;
            LockAspectRatioBtn.ToolTip = UnlockAspectRatioToolTip;
            LockAspectRatioBtn.Image = Utils.Graphics.CreateBitmapSourceFromGdiBitmap(_lockedImage);

            ModifySelectionAspectRatio();
        }

        private void ModifySelectionAspectRatio()
        {
            if (_resizeLab.IsSelecionValid(GetSelection(), false))
            {
                _resizeLab.ChangeShapesAspectRatio(GetSelectedShapes(), IsAspectRatioLocked);
            }
        }

        #endregion

        #region Event Handler: Preview

        private void StretchLeftBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            var resizeAction = new SingleInputResizeAction(shapes => _resizeLab.StretchLeft);

            Preview(selectedShapes, resizeAction, 2);
        }

        private void StretchLeftBtn_MouseLeave(object sender, MouseEventArgs e)
        {
            Reset();
        }

        private void StretchRightBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            var resizeAction = new SingleInputResizeAction(shapes => _resizeLab.StretchRight);

            Preview(selectedShapes, resizeAction, 2);
        }

        private void StretchRightBtn_MouseLeave(object sender, MouseEventArgs e)
        {
            Reset();
        }

        private void StretchTopBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            var resizeAction = new SingleInputResizeAction(shapes => _resizeLab.StretchTop);

            Preview(selectedShapes, resizeAction, 2);
        }

        private void StretchTopBtn_MouseLeave(object sender, MouseEventArgs e)
        {
            Reset();
        }

        private void StretchBottomBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            var resizeAction = new SingleInputResizeAction(shapes => _resizeLab.StretchBottom);

            Preview(selectedShapes, resizeAction, 2);
        }

        private void StretchBottomBtn_MouseLeave(object sender, MouseEventArgs e)
        {
            Reset();
        }

        private void SameWidthBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            var resizeAction = new SingleInputResizeAction(shapes => _resizeLab.ResizeToSameWidth);

            Preview(selectedShapes, resizeAction, 2);
        }

        private void SameWidthBtn_MouseLeave(object sender, MouseEventArgs e)
        {
            Reset();
        }

        private void SameHeightBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            var resizeAction = new SingleInputResizeAction(shapes => _resizeLab.ResizeToSameHeight);

            Preview(selectedShapes, resizeAction, 2);
        }

        private void SameHeightBtn_MouseLeave(object sender, MouseEventArgs e)
        {
            Reset();
        }

        private void SameSizeBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            var resizeAction = new SingleInputResizeAction(shapes => _resizeLab.ResizeToSameHeightAndWidth);

            Preview(selectedShapes, resizeAction, 2);
        }

        private void SameSizeBtn_MouseLeave(object sender, MouseEventArgs e)
        {
            Reset();
        }

        private void FitWidthBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            var resizeAction = new MultiInputResizeAction((shapes, width, height, isAspectRatio) => _resizeLab.FitToWidth);

            Preview(selectedShapes, slideWidth, slideHeight, resizeAction);
        }

        private void FitWidthBtn_MouseLeave(object sender, MouseEventArgs e)
        {
            Reset();
        }

        private void FitHeightBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            var resizeAction = new MultiInputResizeAction((shapes, width, height, isAspectRatio) => _resizeLab.FitToHeight);

            Preview(selectedShapes, slideWidth, slideHeight, resizeAction);
        }

        private void FitHeightBtn_MouseLeave(object sender, MouseEventArgs e)
        {
            Reset();
        }

        private void FillBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            var resizeAction = new MultiInputResizeAction((shapes, width, height, isAspectRatio) => _resizeLab.FitToFill);

            Preview(selectedShapes, slideWidth, slideHeight, resizeAction);
        }

        private void FillBtn_MouseLeave(object sender, MouseEventArgs e)
        {
            Reset();
        }

        #endregion

        #region Helper Functions

        private PowerPoint.ShapeRange GetSelectedShapes(bool handleError = true)
        {
            var selection = GetSelection();

            return _resizeLab.IsSelecionValid(selection, handleError) ? GetSelection().ShapeRange : null;
        }

        private PowerPoint.Selection GetSelection()
        {
            return this.GetCurrentSelection();
        }

        private void CleanOriginalShapes()
        {
            _originalShapeProperties.Clear();
        }

        #endregion
    }
}
