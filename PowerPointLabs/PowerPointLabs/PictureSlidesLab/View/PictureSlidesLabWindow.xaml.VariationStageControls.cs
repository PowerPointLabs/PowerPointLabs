using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.PictureSlidesLab.Model;
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
            UpdateVariantsSliderVisibility();
            UpdatePictureAspectRefreshButtonVisibility();
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
            DisableLoadingStyleOnWindowActivate();
            if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ViewModel.BindSelectedColor(colorDialog.Color,
                    this.GetCurrentSlide().GetNativeSlide(),
                    this.GetCurrentPresentation().SlideWidth,
                    this.GetCurrentPresentation().SlideHeight);
            }
            EnableLoadingStyleOnWindowActivate();
        }

        private void VariantsFontPanel_OnDropDownClosed(object sender, EventArgs e)
        {
            ViewModel.BindSelectedFont(
                this.GetCurrentSlide().GetNativeSlide(),
                this.GetCurrentPresentation().SlideWidth,
                this.GetCurrentPresentation().SlideHeight);
        }

        private void VariantsFontPanel_OnKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Up || e.Key == Key.Down)
            {
                ViewModel.BindSelectedFont(
                    this.GetCurrentSlide().GetNativeSlide(),
                    this.GetCurrentPresentation().SlideWidth,
                    this.GetCurrentPresentation().SlideHeight);
            }
        }

        private void VariantsSlider_OnValueChangedFinal(object sender, EventArgs e)
        {
            var type = e.GetType();
            if (type.Equals(typeof(System.Windows.Input.KeyEventArgs)))
            {
                var eventArgs = (System.Windows.Input.KeyEventArgs)e;
                if (eventArgs.Key != Key.Left && eventArgs.Key != Key.Right)
                {
                    return;
                }
            }

            if (ViewModel.IsSliderValueChanged.Flag)
            {
                ViewModel.BindSelectedSliderValue(
                    this.GetCurrentSlide().GetNativeSlide(),
                    this.GetCurrentPresentation().SlideWidth,
                    this.GetCurrentPresentation().SlideHeight);
                ViewModel.IsSliderValueChanged.Flag = false;
            }
        }

        private void VariantsSlider_OnValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            ViewModel.IsSliderValueChanged.Flag = true;
        }
        #endregion

        #region Helper funcs
        private void UpdateVariantFontPanelVisibility()
        {
            if (VariantsComboBox.SelectedValue == null) return;

            var selectedItem = StylesVariationListBox.SelectedValue as ImageItem;

            var currentCategory = (string) VariantsComboBox.SelectedValue;
            if (currentCategory == TextCollection.PictureSlidesLabText.VariantCategoryFontFamily
                && selectedItem != null
                && selectedItem.Tooltip != TextCollection.PictureSlidesLabText.NoEffect)
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

            var selectedItem = StylesVariationListBox.SelectedValue as ImageItem;

            var currentCategory = (string) VariantsComboBox.SelectedValue;
            if (currentCategory.Contains(TextCollection.PictureSlidesLabText.ColorHasEffect)
                 && selectedItem != null
                 && selectedItem.Tooltip != TextCollection.PictureSlidesLabText.NoEffect)
            {
                VariantsColorPanel.Visibility = Visibility.Visible;
                ViewModel.BindStyleToColorPanel();
            }
            else
            {
                VariantsColorPanel.Visibility = Visibility.Collapsed;
            }
        }

        private void UpdateVariantsSliderVisibility()
        {
            if (VariantsComboBox.SelectedValue == null) return;

            var selectedItem = StylesVariationListBox.SelectedValue as ImageItem;

            var currentCategory = (string)VariantsComboBox.SelectedValue;
            if (IsSliderSupported(currentCategory)
                 && selectedItem != null
                 && selectedItem.Tooltip != TextCollection.PictureSlidesLabText.NoEffect)
            {
                VariantsSlider.Visibility = Visibility.Visible;
                ViewModel.BindStyleToSlider();
            }
            else
            {
                VariantsSlider.Visibility = Visibility.Collapsed;
            }
        }

        private void UpdatePictureAspectRefreshButtonVisibility()
        {
            if (VariantsComboBox.SelectedValue == null) return;

            var currentCategory = (string) VariantsComboBox.SelectedValue;
            if (currentCategory == TextCollection.PictureSlidesLabText.VariantCategoryPicture)
            {
                PictureAspectRefreshButton.Visibility = Visibility.Visible;
            }
            else
            {
                PictureAspectRefreshButton.Visibility = Visibility.Collapsed;
            }
        }

        private Color GetColor(SolidColorBrush brush)
        {
            return Color.FromArgb(brush.Color.A, brush.Color.R, brush.Color.G, brush.Color.B);
        }

        private bool IsSliderSupported(string currentCategory)
        {
            return currentCategory == TextCollection.PictureSlidesLabText.VariantCategoryBlurriness
                 || currentCategory == TextCollection.PictureSlidesLabText.VariantCategoryBrightness
                 || currentCategory.Contains(TextCollection.PictureSlidesLabText.TransparencyHasEffect);
        }
        #endregion
    }
}
