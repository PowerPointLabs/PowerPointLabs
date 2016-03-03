using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows;
using System.Windows.Input;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Utils;
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
            Action<PowerPoint.ShapeRange> resizeAction = shapes => { _resizeLab.StretchLeft(shapes); };

            ModifySelectionAspectRatio();
            ExecuteResizeAction(selectedShape, resizeAction);
        }

        private void StretchRightBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            Action<PowerPoint.ShapeRange> resizeAction = shapes => { _resizeLab.StretchRight(shapes); };

            ModifySelectionAspectRatio();
            ExecuteResizeAction(selectedShapes, resizeAction);
        }

        private void StretchTopBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            Action<PowerPoint.ShapeRange> resizeAction = shapes => { _resizeLab.StretchTop(shapes); };

            ModifySelectionAspectRatio();
            ExecuteResizeAction(selectedShapes, resizeAction);
        }

        private void StretchBottomBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            Action<PowerPoint.ShapeRange> resizeAction = shapes => { _resizeLab.StretchBottom(shapes); };

            ModifySelectionAspectRatio();
            ExecuteResizeAction(selectedShapes, resizeAction);
        }

        #endregion

        #region Event Handler: Same Dimension

        private void SameWidthBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            Action<PowerPoint.ShapeRange> resizeAction = shapes => { _resizeLab.ResizeToSameWidth(shapes); };

            ModifySelectionAspectRatio();
            ExecuteResizeAction(selectedShapes, resizeAction);
        }

        private void SameHeightBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            Action<PowerPoint.ShapeRange> resizeAction = shapes => { _resizeLab.ResizeToSameHeight(shapes); };

            ModifySelectionAspectRatio();
            ExecuteResizeAction(selectedShapes, resizeAction);
        }

        private void SameSizeBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            Action<PowerPoint.ShapeRange> resizeAction = shapes => { _resizeLab.ResizeToSameHeightAndWidth(shapes); };

            ModifySelectionAspectRatio();
            ExecuteResizeAction(selectedShapes, resizeAction);
        }

        #endregion

        #region Event Handler: Fit
        private void FitWidthBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<PowerPoint.ShapeRange, float, float, bool> resizeAction =
                (shapes, referenceWidth, referenceHeight, isAspectRatio) =>
                {
                    _resizeLab.FitToWidth(shapes, referenceWidth, referenceHeight, isAspectRatio);
                };

            ModifySelectionAspectRatio();
            ExecuteResizeAction(selectedShapes, slideWidth, slideHeight, resizeAction);
        }

        private void FitHeightBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<PowerPoint.ShapeRange, float, float, bool> resizeAction =
                (shapes, referenceWidth, referenceHeight, isAspectRatio) =>
                {
                    _resizeLab.FitToHeight(shapes, referenceWidth, referenceHeight, isAspectRatio);
                };

            ModifySelectionAspectRatio();
            ExecuteResizeAction(selectedShapes, slideWidth, slideHeight, resizeAction);
        }

        private void FillBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedShapes = GetSelectedShapes();
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<PowerPoint.ShapeRange, float, float, bool> resizeAction =
                (shapes, referenceWidth, referenceHeight, isAspectRatio) =>
                {
                    _resizeLab.FitToFill(shapes, referenceWidth, referenceHeight, isAspectRatio);
                };

            ModifySelectionAspectRatio();
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
            Action<PowerPoint.ShapeRange> previewAction = shapes => { _resizeLab.StretchLeft(shapes); }; 

            ModifySelectionAspectRatio();
            Preview(selectedShapes, previewAction, 2);
        }

        private void StretchRightBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            Action<PowerPoint.ShapeRange> previewAction = shapes => { _resizeLab.StretchRight(shapes); };

            ModifySelectionAspectRatio();
            Preview(selectedShapes, previewAction, 2);
        }

        private void StretchTopBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            Action<PowerPoint.ShapeRange> previewAction = shapes => { _resizeLab.StretchTop(shapes); };

            ModifySelectionAspectRatio();
            Preview(selectedShapes, previewAction, 2);
        }

        private void StretchBottomBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            Action<PowerPoint.ShapeRange> previewAction = shapes => { _resizeLab.StretchBottom(shapes); };

            ModifySelectionAspectRatio();
            Preview(selectedShapes, previewAction, 2);
        }

        private void SameWidthBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            Action<PowerPoint.ShapeRange> previewAction = shapes => { _resizeLab.ResizeToSameWidth(shapes); };

            ModifySelectionAspectRatio();
            Preview(selectedShapes, previewAction, 2);
        }
        
        private void SameHeightBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            Action<PowerPoint.ShapeRange> previewAction = shapes => { _resizeLab.ResizeToSameHeight(shapes); };

            ModifySelectionAspectRatio();
            Preview(selectedShapes, previewAction, 2);
        }

        private void SameSizeBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            Action<PowerPoint.ShapeRange> previewAction = shapes => { _resizeLab.ResizeToSameHeightAndWidth(shapes); };

            _resizeLab.ChangeShapesAspectRatio(selectedShapes, false);
            Preview(selectedShapes, previewAction, 2);
        }

        private void SameSizeBtn_MouseLeave(object sender, MouseEventArgs e)
        {
            Reset();
            _resizeLab.ChangeShapesAspectRatio(GetSelectedShapes(false), IsAspectRatioLocked);
        }

        private void FitWidthBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<PowerPoint.ShapeRange, float, float, bool> previewAction =
                (shapes, referenceWidth, referenceHeight, isAspectRatio) =>
                {
                    _resizeLab.FitToWidth(shapes, referenceWidth, referenceHeight, isAspectRatio);
                };

            ModifySelectionAspectRatio();
            Preview(selectedShapes, slideWidth, slideHeight, previewAction);
        }

        private void FitHeightBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<PowerPoint.ShapeRange, float, float, bool> previewAction =
                (shapes, referenceWidth, referenceHeight, isAspectRatio) =>
                {
                    _resizeLab.FitToHeight(shapes, referenceWidth, referenceHeight, isAspectRatio);
                };

            ModifySelectionAspectRatio();
            Preview(selectedShapes, slideWidth, slideHeight, previewAction);
        }

        private void FillBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            var selectedShapes = GetSelectedShapes(false);
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;
            Action<PowerPoint.ShapeRange, float, float, bool> previewAction =
                (shapes, referenceWidth, referenceHeight, isAspectRatio) =>
                {
                    _resizeLab.FitToFill(shapes, referenceWidth, referenceHeight, isAspectRatio);
                };

            ModifySelectionAspectRatio();
            Preview(selectedShapes, slideWidth, slideHeight, previewAction);
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

        private void Btn_MouseLeave(object sender, MouseEventArgs e)
        {
            Reset();
        }

        private void SymmetricLeftBtn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void SymmetricLeftBtn_MouseEnter(object sender, MouseEventArgs e)
        {

        }

        private void SymmetricRightBtn_MouseEnter(object sender, MouseEventArgs e)
        {

        }

        private void SymmetricRightBtn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void SymmetricTopBtn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void SymmetricTopBtn_MouseEnter(object sender, MouseEventArgs e)
        {

        }

        private void SymmetricBottomBtn_MouseEnter(object sender, MouseEventArgs e)
        {

        }

        private void SymmetricBottomBtn_Click(object sender, RoutedEventArgs e)
        {

        }

    }
}
