using System.Drawing;
using System.Windows;
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

        public ResizeLabPaneWPF()
        {
            InitializeComponent();
            InitialiseLogicInstance();
            _unlockedImage = new Bitmap(PowerPointLabs.Properties.Resources.ResizeUnlock);
            _lockedImage = new Bitmap(PowerPointLabs.Properties.Resources.ResizeLock);
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
            PowerPoint.ShapeRange selectedShape = GetSelectedShapes();

            if (selectedShape != null)
            {
                _resizeLab.StretchLeft(selectedShape);
            }
        }

        private void StretchRightBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShape = GetSelectedShapes();

            if (selectedShape != null)
            {
                _resizeLab.StretchRight(selectedShape);
            }
        }

        private void StretchTopBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShape = GetSelectedShapes();

            if (selectedShape != null)
            {
                _resizeLab.StretchTop(selectedShape);
            }
        }

        private void StretchBottomBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShape = GetSelectedShapes();

            if (selectedShape != null)
            {
                _resizeLab.StretchBottom(selectedShape);
            }
        }

        #endregion

        #region Event Handler: Same Dimension

        private void SameWidthBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();

            if (selectedShapes != null)
            {
                _resizeLab.ResizeToSameWidth(selectedShapes);
            }
        }

        private void SameHeightBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();

            if (selectedShapes != null)
            {
                _resizeLab.ResizeToSameHeight(selectedShapes);
            }
        }

        private void SameSizeBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();

            if (selectedShapes != null)
            {
                _resizeLab.ResizeToSameHeightAndWidth(selectedShapes);
            }
        }

        #endregion

        #region Event Handler: Fit
        private void FitWidthBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;

            if (selectedShapes != null)
            {
                _resizeLab.FitToWidth(selectedShapes, slideWidth, slideHeight, IsAspectRatioLocked);
            }
        }

        private void FitHeightBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;

            if (selectedShapes != null)
            {
                _resizeLab.FitToHeight(selectedShapes, slideWidth, slideHeight, IsAspectRatioLocked);
            }
        }

        private void FillBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;

            if (selectedShapes != null)
            {
                _resizeLab.FitToFill(selectedShapes, slideWidth, slideHeight, IsAspectRatioLocked);
            }
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

        #region Helper Functions

        private PowerPoint.ShapeRange GetSelectedShapes()
        {
            var selection = GetSelection();
            return _resizeLab.IsSelecionValid(selection, true) ? GetSelection().ShapeRange : null;
        }

        private PowerPoint.Selection GetSelection()
        {
            return this.GetCurrentSelection();
        }
        #endregion

    }
}
