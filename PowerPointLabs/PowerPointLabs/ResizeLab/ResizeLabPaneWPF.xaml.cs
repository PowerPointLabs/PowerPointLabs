using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
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
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;

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
        private Dictionary<string, PowerPoint.Shape> _originalShapes = new Dictionary<string, PowerPoint.Shape>();

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
            var resizeAction = new ResizeAction(shapes => _resizeLab.StretchLeft);

            ExecuteResizeAction(selectedShape, resizeAction);
        }

        private void StretchRightBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShape = GetSelectedShapes();
            var resizeAction = new ResizeAction(shapes => _resizeLab.StretchRight);

            ExecuteResizeAction(selectedShape, resizeAction);
        }

        private void StretchTopBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShape = GetSelectedShapes();
            var resizeAction = new ResizeAction(shapes => _resizeLab.StretchTop);

            ExecuteResizeAction(selectedShape, resizeAction);
        }

        private void StretchBottomBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShape = GetSelectedShapes();
            var resizeAction = new ResizeAction(shapes => _resizeLab.StretchBottom);

            ExecuteResizeAction(selectedShape, resizeAction);
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
                _resizeLab.FitToHight(selectedShapes, slideWidth, slideHeight, IsAspectRatioLocked);
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
            var slideHight = this.GetCurrentPresentation().SlideHeight;
            var slideWidth = this.GetCurrentPresentation().SlideWidth;

            if (selectedShapes != null)
            {
                _resizeLab.RestoreAspectRatio(selectedShapes, slideHight, slideWidth);
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
            var resizeAction = new ResizeAction(shapes => _resizeLab.StretchLeft);

            Preview(selectedShapes, resizeAction);
        }

        private void StretchLeftBtn_MouseLeave(object sender, MouseEventArgs e)
        {
            Reset();
        }

        private void StretchRightBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            var resizeAction = new ResizeAction(shapes => _resizeLab.StretchRight);

            Preview(selectedShapes, resizeAction);
        }

        private void StretchRightBtn_MouseLeave(object sender, MouseEventArgs e)
        {
            Reset();
        }

        #endregion

        #region Helper Functions

        private void ExecuteResizeAction(PowerPoint.ShapeRange selectedShapes, ResizeAction resizeAction)
        {
            if (selectedShapes == null) return;

            var action = resizeAction(selectedShapes);

            Reset();
            action(selectedShapes);
            CleanOriginalShapes();
        }

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
            _originalShapes.Clear();
        }

        private void DoNothing(PowerPoint.ShapeRange selectedShapes) { }


        #endregion
    }
}
