using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using PowerPointLabs.ImagesLab.Model;
using PowerPointLabs.ImagesLab.Util;
using PowerPointLabs.Utils;
using Brush = System.Windows.Media.Brush;
using Color = System.Drawing.Color;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace PowerPointLabs.ImagesLab.View
{
    partial class ImagesLabWindow
    {
        private void CustomizeStyle(IList<StyleOptions> givenStyles = null,
            Dictionary<string, List<StyleVariants>> givenVariants = null)
        {
            UpdateStyleVariationsImagesWhenOpenFlyout(givenStyles, givenVariants);
            OpenVariationsFlyout();
        }

        private void UpdateStyleVariationsImagesWhenOpenFlyout(IList<StyleOptions> givenOptions,
            Dictionary<string, List<StyleVariants>> givenVariants)
        {
            var targetStyleItem = (ImageItem)StylesPreviewListBox.SelectedValue;
            var source = ImageSelectionListBox.SelectedValue as ImageItem;
            
            if (!IsAbleToUpdateVariationsImages(source, targetStyleItem))
            {
                ViewModel.ClearStyleVariationList();
                return;
            }

            ViewModel.InitStyleVariationCategories(givenOptions, givenVariants, targetStyleItem.Tooltip);
            ViewModel.UpdateStyleVariationImages(source);
            SetVariationListBoxSelectedId(0);
            SetVariationListBoxScrollOffset(0);
        }

        public void UpdateStyleVariationsImages(IList<StyleOptions> givenOptions = null,
            Dictionary<string, List<StyleVariants>> givenVariants = null)
        {
            var targetStyleItem = (ImageItem)StylesPreviewListBox.SelectedValue;
            var source = ImageSelectionListBox.SelectedValue as ImageItem;

            if (!IsAbleToUpdateVariationsImages(source, targetStyleItem))
            {
                ViewModel.ClearStyleVariationList();
                return;
            }

            ViewModel.UpdateStyleVariationImages(source);
        }

        private static bool IsAbleToUpdateVariationsImages(ImageItem source, ImageItem targetStyleItem)
        {
            return !(source == null 
                    || source.ImageFile == StoragePath.LoadingImgPath
                    || targetStyleItem == null
                    || targetStyleItem.Tooltip == null
                    || Models.PowerPointCurrentPresentationInfo.CurrentSlide == null);
        }

        private void VariantsComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (StylesVariationListBox.SelectedIndex < 0
                || VariantsComboBox.Items.Count == 0) return;

            ViewModel.UpdateStyleVariationCategories();
        }

        private void ColorPanel_OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            var panel = sender as Border;
            if (panel == null) return;

            var colorDialog = new ColorDialog
            {
                Color = GetColor(panel.Background as SolidColorBrush),
                FullOpen = true
            };
            if (colorDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

            BindColorToStyle(colorDialog.Color);
            BindColorToVariant(colorDialog.Color);
        }

        private Color GetColor(SolidColorBrush brush)
        {
            return Color.FromArgb(brush.Color.A, brush.Color.R, brush.Color.G, brush.Color.B);
        }

        private void BindStyleToColorPanel()
        {
            if (StylesVariationListBox.SelectedIndex < 0
                || VariantsComboBox.Items.Count == 0) return;

            var styleOption = ViewModel.GetStyleVariationStyleOptions(StylesVariationListBox.SelectedIndex);
            var currentCategory = (string)VariantsComboBox.SelectedValue;
            var bc = new BrushConverter();

            if (currentCategory.Contains("Color"))
            {
                var propName = GetPropertyName(currentCategory);
                var type = styleOption.GetType();
                var prop = type.GetProperty(propName);
                var optValue = prop.GetValue(styleOption, null);
                VariantsColorPanel.Background = (Brush)bc.ConvertFrom(optValue);
            }
        }

        private void BindStyleToFontPanel()
        {
            if (StylesVariationListBox.SelectedIndex < 0
                || VariantsComboBox.Items.Count == 0) return;

            var styleOption = ViewModel.GetStyleVariationStyleOptions(StylesVariationListBox.SelectedIndex);
            var currentCategory = (string)VariantsComboBox.SelectedValue;

            if (currentCategory == TextCollection.ImagesLabText.VariantCategoryFontFamily)
            {
                var styleFontFamily = styleOption.GetFontFamily();
                var targetIndex = -1;
                for (var i = 0; i < _fontFamilyList.Count; i++)
                {
                    if (styleFontFamily == _fontFamilyList[i])
                    {
                        targetIndex = i;
                        break;
                    }
                }
                FontPanel.SelectedIndex = targetIndex;
            }
        }

        private void BindColorToStyle(Color color)
        {
            if (StylesVariationListBox.SelectedIndex < 0
                || VariantsComboBox.Items.Count == 0) return;

            var styleOption = ViewModel.GetStyleVariationStyleOptions(StylesVariationListBox.SelectedIndex);
            var currentCategory = (string) VariantsComboBox.SelectedValue;
            var targetColor = StringUtil.GetHexValue(color);

            if (currentCategory.Contains("Color"))
            {
                styleOption.OptionName = "Customized";
                var propName = GetPropertyName(currentCategory);
                var type = styleOption.GetType();
                var prop = type.GetProperty(propName);
                prop.SetValue(styleOption, targetColor, null);
            }
        }

        private void BindColorToVariant(Color color)
        {
            if (StylesVariationListBox.SelectedIndex < 0
                || VariantsComboBox.Items.Count == 0) return;

            var currentCategory = (string)VariantsComboBox.SelectedValue;
            var styleVariant = ViewModel.GetStyleVariationStyleVariants(currentCategory, 
                StylesVariationListBox.SelectedIndex);

            if (currentCategory.Contains("Color"))
            {
                styleVariant.Set("OptionName", "Customized");
                styleVariant.Set(GetPropertyName(currentCategory), StringUtil.GetHexValue(color));
            }
        }

        private void BindFontToStyle(string font)
        {
            if (StylesVariationListBox.SelectedIndex < 0
                || VariantsComboBox.Items.Count == 0) return;

            var styleOption = ViewModel.GetStyleVariationStyleOptions(StylesVariationListBox.SelectedIndex);
            var currentCategory = (string)VariantsComboBox.SelectedValue;

            if (currentCategory == TextCollection.ImagesLabText.VariantCategoryFontFamily)
            {
                styleOption.OptionName = "Customized";
                styleOption.FontFamily = font;
            }
        }

        private void BindFontToVariant(string font)
        {
            if (StylesVariationListBox.SelectedIndex < 0
                || VariantsComboBox.Items.Count == 0) return;

            var currentCategory = (string)VariantsComboBox.SelectedValue;
            var styleVariant = ViewModel.GetStyleVariationStyleVariants(currentCategory,
                StylesVariationListBox.SelectedIndex);

            if (currentCategory == TextCollection.ImagesLabText.VariantCategoryFontFamily)
            {
                styleVariant.Set("OptionName", "Customized");
                styleVariant.Set("FontFamily", font);
            }
        }

        private void VariationListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (StylesPreviewListBox.SelectedValue == null) return;

            if (StylesVariationListBox.SelectedValue == null)
            {
                StyleApplyButton.IsEnabled = false;
            }
            else
            {
                StyleApplyButton.IsEnabled = true;

                var targetStyleOption = ViewModel
                    .GetStyleVariationStyleOptions(StylesVariationListBox.SelectedIndex);
                ViewModel.SetStyleDesignerOptions(targetStyleOption);

                var currentCategory = (string)VariantsComboBox.SelectedValue;
                
                if (currentCategory.Contains("Color"))
                {
                    VariantsColorPanel.Visibility = Visibility.Visible;
                    BindStyleToColorPanel();
                }
                else
                {
                    VariantsColorPanel.Visibility = Visibility.Collapsed;
                }

                if (currentCategory == TextCollection.ImagesLabText.VariantCategoryFontFamily)
                {
                    FontPanel.Visibility = Visibility.Visible;
                    BindStyleToFontPanel();
                }
                else
                {
                    FontPanel.Visibility = Visibility.Collapsed;
                }
            }
        }

        private void StyleApplyButton_OnClick(object sender, RoutedEventArgs e)
        {
            var image = (ImageItem) ImageSelectionListBox.SelectedValue;
            if (image == null) return;
            ViewModel.ApplyStyleInVariationStage(image);
        }

        private void VariationFlyoutBackButton_OnClick(object sender, RoutedEventArgs e)
        {
            CloseVariationsFlyout();
        }

        private void CloseVariationsFlyout()
        {
            if (!_isVariationsFlyoutOpen) return;

            var right2LeftToHideTranslate = new TranslateTransform();
            StyleVariationsFlyout.RenderTransform = right2LeftToHideTranslate;
            var right2LeftToHideAnimation = new DoubleAnimation(0, -StyleVariationsFlyout.ActualWidth,
                TimeSpan.FromMilliseconds(350))
            {
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut },
                AccelerationRatio = 0.5
            };
            right2LeftToHideAnimation.Completed += (sender, args) =>
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    StyleVariationsFlyout.Visibility = Visibility.Collapsed;
                    if (_latestImageChangedTime > _latestPreviewUpdateTime)
                    {
                        UpdatePreviewImages();
                    }
                }));
            };

            right2LeftToHideTranslate.BeginAnimation(TranslateTransform.XProperty, right2LeftToHideAnimation);
            _isVariationsFlyoutOpen = false;
        }

        private void OpenVariationsFlyout()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_isVariationsFlyoutOpen) return;

                var left2RightToShowTranslate = new TranslateTransform {X = -StylesPreviewGrid.ActualWidth};
                StyleVariationsFlyout.RenderTransform = left2RightToShowTranslate;
                StyleVariationsFlyout.Visibility = Visibility.Visible;
                var left2RightToShowAnimation = new DoubleAnimation(-StylesPreviewGrid.ActualWidth, 0,
                    TimeSpan.FromMilliseconds(350))
                {
                    EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut },
                    AccelerationRatio = 0.5
                };

                left2RightToShowTranslate.BeginAnimation(TranslateTransform.XProperty, left2RightToShowAnimation);
                _isVariationsFlyoutOpen = true;
            }));
        }

        private void FontPanel_OnDropDownClosed(object sender, EventArgs e)
        {
            BindNewlySelectedFont();
        }

        private void FontPanel_OnKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Up || e.Key == Key.Down)
            {
                BindNewlySelectedFont();
            }
        }

        private void BindNewlySelectedFont()
        {
            var selectedFontFamily = (FontFamily) FontPanel.SelectedValue;
            BindFontToStyle(selectedFontFamily.Source);
            BindFontToVariant(selectedFontFamily.Source);
            UpdatePreviewImages();
        }

        private void InitFontFamilyList()
        {
            var fonts = Fonts.SystemFontFamilies;
            foreach (var font in fonts)
            {
                _fontFamilyList.Add(font.Source);
            }
        }

        private string GetPropertyName(string categoryName)
        {
            return categoryName.Replace(" ", string.Empty);
        }
    }
}
