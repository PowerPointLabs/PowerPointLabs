﻿using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.TextCollection;

using Color = System.Drawing.Color;
using Forms = System.Windows.Forms;

namespace PowerPointLabs.PictureSlidesLab.Views
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
            Border panel = sender as Border;
            if (panel == null)
            {
                return;
            }

            ColorDialog colorDialog = new ColorDialog
            {
                Color = GetColor(panel.Background as SolidColorBrush),
                FullOpen = true
            };
            DisableLoadingStyleOnWindowActivate();
            if (colorDialog.ShowDialog() == Forms.DialogResult.OK)
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
            Type type = e.GetType();
            if (type.Equals(typeof(System.Windows.Input.KeyEventArgs)))
            {
                System.Windows.Input.KeyEventArgs eventArgs = (System.Windows.Input.KeyEventArgs)e;
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
            if (VariantsComboBox.SelectedValue == null)
            {
                return;
            }

            ImageItem selectedItem = StylesVariationListBox.SelectedValue as ImageItem;

            string currentCategory = (string) VariantsComboBox.SelectedValue;
            if (currentCategory == PictureSlidesLabText.VariantCategoryFontFamily
                && selectedItem != null
                && selectedItem.Tooltip != PictureSlidesLabText.NoEffect)
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
            if (VariantsComboBox.SelectedValue == null)
            {
                return;
            }

            ImageItem selectedItem = StylesVariationListBox.SelectedValue as ImageItem;

            string currentCategory = (string) VariantsComboBox.SelectedValue;
            if (currentCategory.Contains(PictureSlidesLabText.ColorHasEffect)
                 && selectedItem != null
                 && selectedItem.Tooltip != PictureSlidesLabText.NoEffect)
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
            if (VariantsComboBox.SelectedValue == null)
            {
                return;
            }

            ImageItem selectedItem = StylesVariationListBox.SelectedValue as ImageItem;

            string currentCategory = (string)VariantsComboBox.SelectedValue;
            if (IsSliderSupported(currentCategory)
                 && selectedItem != null
                 && selectedItem.Tooltip != PictureSlidesLabText.NoEffect)
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
            if (VariantsComboBox.SelectedValue == null)
            {
                return;
            }

            string currentCategory = (string) VariantsComboBox.SelectedValue;
            if (currentCategory == PictureSlidesLabText.VariantCategoryPicture)
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
            return currentCategory == PictureSlidesLabText.VariantCategoryBlurriness
                 || currentCategory == PictureSlidesLabText.VariantCategoryBrightness
                 || currentCategory.Contains(PictureSlidesLabText.TransparencyHasEffect);
        }
        #endregion
    }
}
