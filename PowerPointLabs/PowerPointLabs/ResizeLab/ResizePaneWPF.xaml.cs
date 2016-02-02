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
    public partial class ResizePaneWPF : UserControl
    {
        public static bool IsAspectRatioLocked { get; set; }
        private const string UnlockText = "Unlock";
        private const string LockText = "Lock";

        public ResizePaneWPF()
        {
            InitializeComponent();
            UnlockAspectRatio();
        }

        #region Event Handler: Strech and Shrink

        private void StretchLeftBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShape = GetSelectedShapes();

            if (selectedShape != null)
            {
                ResizeLabMain.StretchLeft(selectedShape);
            }
        }

        private void StretchRightBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShape = GetSelectedShapes();

            if (selectedShape != null)
            {
                ResizeLabMain.StretchRight(selectedShape);
            }
        }

        private void StretchTopBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShape = GetSelectedShapes();

            if (selectedShape != null)
            {
                ResizeLabMain.StretchTop(selectedShape);
            }
        }

        private void StretchBottomBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShape = GetSelectedShapes();

            if (selectedShape != null)
            {
                ResizeLabMain.StretchBottom(selectedShape);
            }
        }

        #endregion

        #region Event Handler: Same Dimension

        private void SameWidthBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();

            if (selectedShapes != null)
            {
                ResizeLabMain.ResizeToSameWidth(selectedShapes);
            }
        }

        private void SameHeightBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();

            if (selectedShapes != null)
            {
                ResizeLabMain.ResizeToSameHeight(selectedShapes);
            }
        }

        private void SameSizeBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();

            if (selectedShapes != null)
            {
                ResizeLabMain.ResizeToSameHeightAndWidth(selectedShapes);
            }
        }

        #endregion

        #region Event Handler: Fit
        private void FitWidthBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();

            if (selectedShapes != null)
            {
                ResizeLabMain.FitToWidth(selectedShapes, IsAspectRatioLocked);
            }
        }

        private void FitHeightBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();

            if (selectedShapes != null)
            {
                ResizeLabMain.FitToHight(selectedShapes, IsAspectRatioLocked);
            }
        }

        private void FillBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();

            if (selectedShapes != null)
            {
                ResizeLabMain.FitToFill(selectedShapes);
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
            if (ResizeLabMain.IsShapeSelection(GetSelection()))
            {
                ResizeLabMain.ChangeShapesAspectRatio(GetSelectedShapes(), IsAspectRatioLocked);
            }
        }

        #endregion

        #region Helper Functions

        private PowerPoint.ShapeRange GetSelectedShapes()
        {
            var selection = GetSelection();
            return ResizeLabMain.IsSelecionValid(selection) ? GetSelection().ShapeRange : null;
        }

        private PowerPoint.Selection GetSelection()
        {
            return PowerPointCurrentPresentationInfo.CurrentSelection;
        }
        #endregion

    }
}
