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
        public ResizePaneWPF()
        {
            InitializeComponent();
        }

        #region Event Handler: Strech and Shrink

        private void StretchLeftBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!ResizeLabMain.IsSelecionValid(GetSelection()))
            {
                return;
            }
            ResizeLabMain.StretchLeft(GetSelectedShapes());
        }

        private void StretchRightBtn_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private void StretchTopBtn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void StretchBottomBtn_Click(object sender, RoutedEventArgs e)
        {

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
                ResizeLabMain.FitToWidth(selectedShapes);
            }
        }

        private void FitHeightBtn_Click(object sender, RoutedEventArgs e)
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes();

            if (selectedShapes != null)
            {
                ResizeLabMain.FitToHight(selectedShapes);
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

        }

        private void RestoreAspectRatioBtn_Click(object sender, RoutedEventArgs e)
        {

        }

        #endregion

        #region Helper Functions

        private static PowerPoint.ShapeRange GetSelectedShapes()
        {
            PowerPoint.Selection selection = GetSelection();

            if (ResizeLabMain.IsSelecionValid(selection))
            {
                return GetSelection().ShapeRange;
            }
            return null;
        }

        private static PowerPoint.Selection GetSelection()
        {
            return PowerPointCurrentPresentationInfo.CurrentSelection;
        }
        #endregion

    }
}
