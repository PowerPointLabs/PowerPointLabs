using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using Color = System.Drawing.Color;

namespace PowerPointLabs.PictureSlidesLab.View
{
    partial class PictureSlidesLabWindow
    {
        ///////////////////////////////////////////////////////////////
        // Implemented variation stage controls in UI
        ///////////////////////////////////////////////////////////////

        #region Visibility controls and binding style option to property
        /// <summary>
        /// update controls visibility here
        /// </summary>
        private void UpdateVariationStageControls()
        {
            UpdateVariantsColorPanelVisibility();
            UpdateVariantFontPanelVisibility();
        }
        #endregion

        #region Binding selected property
        /// <summary>
        /// open a color dialog to customize color when user clicks the color panel in the variation stage
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void VariantsColorPanel_OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            var panel = sender as Border;
            if (panel == null) return;

            var colorDialog = new ColorDialog
            {
                Color = GetColor(panel.Background as SolidColorBrush),
                FullOpen = true
            };
            if (colorDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

            ViewModel.BindSelectedColor(colorDialog.Color);
        }

        private void VariantsFontPanel_OnDropDownClosed(object sender, EventArgs e)
        {
            ViewModel.BindSelectedFont();
        }

        private void VariantsFontPanel_OnKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Up || e.Key == Key.Down)
            {
                ViewModel.BindSelectedFont();
            }
        }
        #endregion

        #region Helper funcs
        private void UpdateVariantFontPanelVisibility()
        {
            if (VariantsComboBox.SelectedValue == null) return;

            var currentCategory = (string) VariantsComboBox.SelectedValue;
            if (currentCategory == TextCollection.PictureSlidesLabText.VariantCategoryFontFamily)
            {
                FontPanel.Visibility = Visibility.Visible;
                ViewModel.BindStyleToFontPanel();
            }
            else
            {
                FontPanel.Visibility = Visibility.Collapsed;
            }
        }

        private void UpdateVariantsColorPanelVisibility()
        {
            if (VariantsComboBox.SelectedValue == null) return;

            var currentCategory = (string) VariantsComboBox.SelectedValue;
            if (currentCategory.Contains("Color"))
            {
                VariantsColorPanel.Visibility = Visibility.Visible;
                ViewModel.BindStyleToColorPanel();
            }
            else
            {
                VariantsColorPanel.Visibility = Visibility.Collapsed;
            }
        }

        private Color GetColor(SolidColorBrush brush)
        {
            return Color.FromArgb(brush.Color.A, brush.Color.R, brush.Color.G, brush.Color.B);
        }
        #endregion
    }
}
