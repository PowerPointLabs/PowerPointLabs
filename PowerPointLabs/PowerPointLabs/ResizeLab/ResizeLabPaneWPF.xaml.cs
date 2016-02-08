using System;
using System.Collections.Generic;
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
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// Interaction logic for ResizePane.xaml
    /// </summary>
    public partial class ResizeLabPaneWPF : IResizeLabPane
    {
        private ResizeLabMain _resizeLab;
        public static bool IsAspectRatioLocked { get; set; }
        private const string UnlockText = "Unlock";
        private const string LockText = "Lock";

        public ResizeLabPaneWPF()
        {
            InitializeComponent();
            InitialiseLogicInstance();
            UnlockAspectRatio();
        }

        internal void InitialiseLogicInstance()
        {
            if (_resizeLab == null)
            {
                _resizeLab = new ResizeLabMain(this);
            }
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

            if (selectedShapes != null)
            {
                _resizeLab.FitToWidth(selectedShapes, IsAspectRatioLocked);
            }
        }

        private void FitHeightBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();

            if (selectedShapes != null)
            {
                _resizeLab.FitToHight(selectedShapes, IsAspectRatioLocked);
            }
        }

        private void FillBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();

            if (selectedShapes != null)
            {
                _resizeLab.FitToFill(selectedShapes, IsAspectRatioLocked);
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
            var slideHight = PowerPointPresentation.Current.SlideHeight;
            var slideWidth = PowerPointPresentation.Current.SlideWidth;

            if (selectedShapes != null)
            {
                _resizeLab.RestoreAspectRatio(selectedShapes, slideHight, slideWidth);
            }
        }

        private void UnlockAspectRatio()
        {
            IsAspectRatioLocked = false;
            LockAspectRatioBtn.Text = UnlockText;

            ModifySelectionAspectRatio();
        }

        private void LockAspectRatio()
        {
            IsAspectRatioLocked = true;
            LockAspectRatioBtn.Text = LockText;

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
            return PowerPointCurrentPresentationInfo.CurrentSelection;
        }
        #endregion

    }
}
