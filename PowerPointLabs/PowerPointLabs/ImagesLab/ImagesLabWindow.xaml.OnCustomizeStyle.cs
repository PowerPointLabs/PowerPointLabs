using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Forms;
using System.Windows.Media.Animation;
using PowerPointLabs.ImagesLab.Domain;
using PowerPointLabs.ImagesLab.Factory;
using PowerPointLabs.ImagesLab.Util;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Exceptions;
using Brush = System.Windows.Media.Brush;
using Color = System.Drawing.Color;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace PowerPointLabs.ImagesLab
{
    partial class ImagesLabWindow
    {
        // list that holds font families
        private readonly List<string> _fontFamilyList = new List<string>();

        private bool _isVariationsFlyoutOpen;

        private string _previousVariantsCategory;
        private IList<StyleOptions> _styleOptions;
        private Dictionary<string, List<StyleVariants>> _styleVariants;

        private void UpdateStyleVariationsImages(bool isOpenFlyout = false, IList<StyleOptions> givenOptions = null,
            Dictionary<string, List<StyleVariants>> givenVariants = null)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (StylesPreviewListBox.SelectedValue == null) return;

                var targetStyle = (ImageItem) StylesPreviewListBox.SelectedValue;
                var source = ImageSelectionListBox.SelectedValue as ImageItem;

                if (source == null || source.ImageFile == StoragePath.LoadingImgPath
                    || Models.PowerPointCurrentPresentationInfo.CurrentSlide == null)
                {
                    StylesVariationList.Clear();
                    return;
                }

                Assumption.Made(targetStyle != null, "target style is null/empty");

                try
                {
                    Double scrollOffset = 0f;
                    var scrollViewer = ListBoxUtil.FindScrollViewer(StylesVariationListBox);
                    if (scrollViewer != null)
                    {
                        scrollOffset = scrollViewer.VerticalOffset;
                    }
                    var selectedId = StylesVariationListBox.SelectedIndex >= 0 ? StylesVariationListBox.SelectedIndex : 0;
                    StylesVariationList.Clear();

                    if (isOpenFlyout)
                    {
                        scrollOffset = 0;
                        selectedId = 0;
                        InitStylesVariationFlyout(givenOptions, givenVariants, targetStyle);
                    }

                    foreach (var styleOption in _styleOptions)
                    {
                        UpdateStyleVariationsImage(styleOption, source);
                    }

                    StylesVariationListBox.SelectedIndex = selectedId;
                    if (scrollViewer != null)
                    {
                        scrollViewer.ScrollToVerticalOffset(scrollOffset);
                    }
                }
                catch
                {
                    ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorImageCorrupted);
                }
            }));
        }

        private void InitStylesVariationFlyout(IList<StyleOptions> givenOptions, 
            Dictionary<string, List<StyleVariants>> givenVariants, ImageItem targetStyle)
        {
            _styleOptions = givenOptions ?? StyleOptionsFactory.GetStylesVariationOptions(targetStyle.Tooltip);
            _styleVariants = givenVariants ?? StyleVariantsFactory.GetVariants(targetStyle.Tooltip);

            VariantsComboBox.Items.Clear();
            foreach (var key in _styleVariants.Keys)
            {
                VariantsComboBox.Items.Add(key);
            }
            VariantsComboBox.SelectedIndex = 0;
            _previousVariantsCategory = (string) VariantsComboBox.SelectedValue;

            // default style options (in preview stage)
            var defaultStyleOptions = StyleOptionsFactory.GetStylesPreviewOption(targetStyle.Tooltip);
            var currentVariants = _styleVariants.Values.First();
            var variantIndexWithoutEffect = -1;
            for (var i = 0; i < currentVariants.Count; i++)
            {
                if (currentVariants[i].IsNoEffect(defaultStyleOptions))
                {
                    variantIndexWithoutEffect = i;
                    break;
                }
            }

            // swap the no-effect variant with the current selected style's corresponding variant
            // so that to achieve continuity.
            // in order to swap, style option provided from StyleOptionsFactory should have
            // corresponding values specified in StyleVariantsFactory. e.g., an option generated
            // from factory has overlay transparency of 35, then in order to swap, it should have
            // a variant of overlay transparency of 35. Otherwise it cannot swap, because variants
            // don't match any values in the style options.
            if (variantIndexWithoutEffect != -1 && givenOptions == null)
            {
                // swap style variant
                var tempVariant = currentVariants[variantIndexWithoutEffect];
                currentVariants[variantIndexWithoutEffect] =
                    currentVariants[0];
                currentVariants[0] = tempVariant;
                // swap default style options (in variation stage)
                var tempStyleOpt = _styleOptions[variantIndexWithoutEffect];
                _styleOptions[variantIndexWithoutEffect] =
                    _styleOptions[0];
                _styleOptions[0] = tempStyleOpt;
            }

            for (var i = 0; i < currentVariants.Count && i < _styleOptions.Count; i++)
            {
                currentVariants[i].Apply(_styleOptions[i]);
            }
        }

        private void VariantsComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (StylesVariationListBox.SelectedIndex < 0
                || VariantsComboBox.Items.Count == 0) return;

            var targetVariants = _styleVariants[_previousVariantsCategory];
            if (targetVariants.Count == 0) return;

            var targetVariationSelectedIndex = StylesVariationListBox.SelectedIndex > 0 && 
                StylesVariationListBox.SelectedIndex < targetVariants.Count
                ? StylesVariationListBox.SelectedIndex
                : 0;
            var targetVariant = targetVariants[targetVariationSelectedIndex];
            foreach (var option in _styleOptions)
            {
                targetVariant.Apply(option);
            }

            var currentVariantsCategory = (string) VariantsComboBox.SelectedValue;
            if (currentVariantsCategory != TextCollection.ImagesLabText.VariantCategoryFontColor
                && _previousVariantsCategory != TextCollection.ImagesLabText.VariantCategoryFontColor)
            {
                // apply font color variant,
                // because default styles may contain special font color settings, but not in variants
                var fontColorVariant = new StyleVariants(new Dictionary<string, object>
                {
                    {"FontColor", _styleOptions[targetVariationSelectedIndex].FontColor}
                });
                foreach (var option in _styleOptions)
                {
                    fontColorVariant.Apply(option);
                }
            }

            var nextCategoryVariants = _styleVariants[currentVariantsCategory];
            var variantIndexWithoutEffect = -1;
            for (var i = 0; i < nextCategoryVariants.Count; i++)
            {
                if (nextCategoryVariants[i].IsNoEffect(_styleOptions[targetVariationSelectedIndex]))
                {
                    variantIndexWithoutEffect = i;
                    break;
                }
            }
            // swap the no-effect variant with the current selected style's corresponding variant
            // so that to achieve an effect: jumpt between different category wont change the
            // selected style
            if (variantIndexWithoutEffect != -1)
            {
                var temp = nextCategoryVariants[variantIndexWithoutEffect];
                nextCategoryVariants[variantIndexWithoutEffect] =
                    nextCategoryVariants[targetVariationSelectedIndex];
                nextCategoryVariants[targetVariationSelectedIndex] = temp;
            }

            for (var i = 0; i < nextCategoryVariants.Count && i < _styleOptions.Count; i++)
            {
                nextCategoryVariants[i].Apply(_styleOptions[i]);
            }

            _previousVariantsCategory = currentVariantsCategory;
            UpdateStyleVariationsImages();
        }

        private void UpdateStyleVariationsImage(StyleOptions opt, ImageItem source)
        {
            PreviewPresentation.SetStyleOptions(opt);
            var previewInfo = PreviewPresentation.PreviewApplyStyle(source);
            StylesVariationList.Add(new ImageItem
            {
                ImageFile = previewInfo.PreviewApplyStyleImagePath,
                Tooltip = opt.OptionName
            });
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

            var styleOption = _styleOptions[StylesVariationListBox.SelectedIndex];
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

            var styleOption = _styleOptions[StylesVariationListBox.SelectedIndex];
            var currentCategory = (string)VariantsComboBox.SelectedValue;

            if (currentCategory == TextCollection.ImagesLabText.VariantCategoryFontFamily)
            {
                var styleFontFamily = styleOption.GetFontFamily();
                var targetIndex = -1;
                for(var i = 0; i < _fontFamilyList.Count; i++)
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

            var styleOption = _styleOptions[StylesVariationListBox.SelectedIndex];
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
            var styleVariant = _styleVariants[currentCategory][StylesVariationListBox.SelectedIndex];

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

            var styleOption = _styleOptions[StylesVariationListBox.SelectedIndex];
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
            var styleVariant = _styleVariants[currentCategory][StylesVariationListBox.SelectedIndex];

            if (currentCategory == TextCollection.ImagesLabText.VariantCategoryFontFamily)
            {
                styleVariant.Set("OptionName", "Customized");
                styleVariant.Set("FontFamily", font);
            }
        }

        // TODO split files into APIs and helper functions
        private void CustomizeStyle(IList<StyleOptions> givenStyles = null, 
            Dictionary<string, List<StyleVariants>> givenVariants = null)
        {
            var source = (ImageItem)ImageSelectionListBox.SelectedValue;
            var targetStyle = StylesPreviewListBox.SelectedItems;
            if (source == null || targetStyle == null || targetStyle.Count == 0) return;

            UpdateStyleVariationsImages(isOpenFlyout: true, givenOptions: givenStyles, givenVariants: givenVariants);
            OpenVariationsFlyout();
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

                var targetStyleOption = _styleOptions[StylesVariationListBox.SelectedIndex];
                PreviewPresentation.SetStyleOptions(targetStyleOption);

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
            ApplyStyleInVariationStage();
        }

        private void ApplyStyleInVariationStage()
        {
            if (StylesPreviewListBox.SelectedValue == null) return;

            var source = ImageSelectionListBox.SelectedValue as ImageItem;
            Assumption.Made(source != null, "source item is null/empty");

            try
            {
                PreviewPresentation.ApplyStyle(source);
                OpenSuccessfullyAppliedDialog();
            }
            catch (AssumptionFailedException)
            {
                ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorNoSelectedSlide);
            }
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
                        DoPreview();
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
            DoPreview();
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
