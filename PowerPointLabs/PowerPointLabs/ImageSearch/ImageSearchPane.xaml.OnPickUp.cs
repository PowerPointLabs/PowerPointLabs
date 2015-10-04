using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Forms;
using System.Windows.Media.Animation;
using MahApps.Metro.Controls.Dialogs;
using PowerPointLabs.AutoUpdate;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Exceptions;
using Brush = System.Windows.Media.Brush;
using Color = System.Drawing.Color;

namespace PowerPointLabs.ImageSearch
{
    partial class ImageSearchPane
    {
        private bool _isVariationsFlyoutOpen;

        private string _previousVariantsCategory;
        private Dictionary<string, int> _selectedVariants;
        private IList<StyleOptions> _styleOptions;
        private Dictionary<string, List<StyleVariants>> _styleVariants; 

        private void UpdateStyleVariationsImages(bool isOpenFlyout = false)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (PreviewListBox.SelectedValue == null) return;

                var targetStyle = (ImageItem) PreviewListBox.SelectedValue;
                var source = SearchListBox.SelectedValue as ImageItem;
                Assumption.Made(source != null && targetStyle != null, "source item or target style is null/empty");

                try
                {
                    var selectedId = VariationListBox.SelectedIndex >= 0 ? VariationListBox.SelectedIndex : 0;
                    VariationList.Clear();

                    if (isOpenFlyout)
                    {
                        selectedId = 0;
                        _styleOptions = StyleOptionsFactory.GetOptions(targetStyle.Tooltip);
                        _styleVariants = StyleVariantsFactory.GetVariants(targetStyle.Tooltip);
                        _selectedVariants = new Dictionary<string, int>();

                        VariantsComboBox.Items.Clear();
                        foreach (var key in _styleVariants.Keys)
                        {
                            VariantsComboBox.Items.Add(key);
                            _selectedVariants.Add(key, 0);
                        }
                        VariantsComboBox.SelectedIndex = 0;
                        _previousVariantsCategory = (string) VariantsComboBox.SelectedValue;

                        foreach (var variants in _styleVariants.Values)
                        {
                            for (var i = 0; i < variants.Count && i < _styleOptions.Count; i++)
                            {
                                variants[i].Apply(_styleOptions[i]);
                            }
                            break;
                        }
                    }

                    foreach (var styleOption in _styleOptions)
                    {
                        UpdateStyleVariationsImage(styleOption, source);
                    }

                    VariationListBox.SelectedIndex = selectedId;
                    if (source.FullSizeImageFile != null)
                    {
                        SetProgressingRingStatus(false);
                    }
                    VariationListBox.ScrollIntoView(VariationListBox.SelectedItem);
                }
                catch
                {
                    // ignore, selected slide may be null
                }
            }));
        }

        private void VariantsComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (VariationListBox.SelectedIndex < 0
                || VariantsComboBox.Items.Count == 0) return;

            _selectedVariants[_previousVariantsCategory] = VariationListBox.SelectedIndex;

            var targetVariants = _styleVariants[_previousVariantsCategory];
            if (targetVariants.Count == 0) return;

            var targetVariationSelectedIndex = VariationListBox.SelectedIndex > 0 && 
                VariationListBox.SelectedIndex < targetVariants.Count
                ? VariationListBox.SelectedIndex
                : 0;
            var targetVariant = targetVariants[targetVariationSelectedIndex];
            foreach (var option in _styleOptions)
            {
                targetVariant.Apply(option);
            }

            var currentVariantsCategory = (string) VariantsComboBox.SelectedValue;
            if (currentVariantsCategory != TextCollection.ImagesLabText.VariantCategoryTextColor
                && _previousVariantsCategory != TextCollection.ImagesLabText.VariantCategoryTextColor)
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
            for (var i = 0; i < nextCategoryVariants.Count && i < _styleOptions.Count; i++)
            {
                nextCategoryVariants[i].Apply(_styleOptions[i]);
            }

            _previousVariantsCategory = currentVariantsCategory;
            VariationListBox.SelectedIndex = _selectedVariants[currentVariantsCategory];
            UpdateStyleVariationsImages();
        }

        private void UpdateStyleVariationsImage(StyleOptions opt, ImageItem source)
        {
            PreviewPresentation.SetStyleOptions(opt);
            var previewInfo = PreviewPresentation.PreviewApplyStyle(source);
            VariationList.Add(new ImageItem
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
            if (VariationListBox.SelectedIndex < 0
                || VariantsComboBox.Items.Count == 0) return;

            var styleOption = _styleOptions[VariationListBox.SelectedIndex];
            var currentCategory = (string)VariantsComboBox.SelectedValue;
            var bc = new BrushConverter();

            if (currentCategory.Contains("Color"))
            {
                switch (currentCategory)
                {
                    case TextCollection.ImagesLabText.VariantCategoryTextColor:
                        VariantsColorPanel.Background = (Brush) bc.ConvertFrom(styleOption.FontColor);
                        break;
                    case TextCollection.ImagesLabText.VariantCategoryTextBoxColor:
                        VariantsColorPanel.Background = (Brush) bc.ConvertFrom(styleOption.TextBoxOverlayColor);
                        break;
                    case TextCollection.ImagesLabText.VariantCategoryBannerColor:
                        VariantsColorPanel.Background = (Brush) bc.ConvertFrom(styleOption.BannerOverlayColor);
                        break;
                    case TextCollection.ImagesLabText.VariantCategoryOverlayColor:
                        VariantsColorPanel.Background = (Brush) bc.ConvertFrom(styleOption.OverlayColor);
                        break;
                }
            }
        }

        private void BindColorToStyle(Color color)
        {
            if (VariationListBox.SelectedIndex < 0
                || VariantsComboBox.Items.Count == 0) return;

            var styleOption = _styleOptions[VariationListBox.SelectedIndex];
            var currentCategory = (string) VariantsComboBox.SelectedValue;
            var targetColor = StringUtil.GetHexValue(color);

            if (currentCategory.Contains("Color"))
            {
                switch (currentCategory)
                {
                    case TextCollection.ImagesLabText.VariantCategoryTextColor:
                        styleOption.FontColor = targetColor;
                        break;
                    case TextCollection.ImagesLabText.VariantCategoryTextBoxColor:
                        styleOption.TextBoxOverlayColor = targetColor;
                        break;
                    case TextCollection.ImagesLabText.VariantCategoryBannerColor:
                        styleOption.BannerOverlayColor = targetColor;
                        break;
                    case TextCollection.ImagesLabText.VariantCategoryOverlayColor:
                        styleOption.OverlayColor = targetColor;
                        break;
                }
            }
        }

        private void BindColorToVariant(Color color)
        {
            if (VariationListBox.SelectedIndex < 0
                || VariantsComboBox.Items.Count == 0) return;

            var currentCategory = (string)VariantsComboBox.SelectedValue;
            var styleVariant = _styleVariants[currentCategory][VariationListBox.SelectedIndex];

            if (currentCategory.Contains("Color"))
            {
                switch (currentCategory)
                {
                    case TextCollection.ImagesLabText.VariantCategoryTextColor:
                        styleVariant.Set("FontColor", StringUtil.GetHexValue(color));
                        break;
                    case TextCollection.ImagesLabText.VariantCategoryTextBoxColor:
                        styleVariant.Set("TextBoxOverlayColor", StringUtil.GetHexValue(color));
                        break;
                    case TextCollection.ImagesLabText.VariantCategoryBannerColor:
                        styleVariant.Set("BannerOverlayColor", StringUtil.GetHexValue(color));
                        break;
                    case TextCollection.ImagesLabText.VariantCategoryOverlayColor:
                        styleVariant.Set("OverlayColor", StringUtil.GetHexValue(color));
                        break;
                }
            }
        }

        private void PickUpStyle()
        {
            var source = (ImageItem)SearchListBox.SelectedValue;
            var targetStyle = PreviewListBox.SelectedItems;
            if (source == null || targetStyle == null || targetStyle.Count == 0) return;

            UpdateStyleVariationsImages(isOpenFlyout: true);
            OpenVariationsFlyout();
        }

        private void VariationListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PreviewListBox.SelectedValue == null) return;

            if (VariationListBox.SelectedValue == null)
            {
                StyleApplyButton.IsEnabled = false;
                StyleCustomizeButton.IsEnabled = false;
            }
            else
            {
                StyleApplyButton.IsEnabled = true;
                StyleCustomizeButton.IsEnabled = true;

                var targetStyleOption = _styleOptions[VariationListBox.SelectedIndex];
                targetStyleOption.PropertyChanged += (o, args) =>
                {
                    _latestStyleOptionsUpdateTime = DateTime.Now;
                };

                PreviewPresentation.SetStyleOptions(targetStyleOption);
                OptionsPane2.DataContext = targetStyleOption;

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
            }
        }

        private void StyleApplyButton_OnClick(object sender, RoutedEventArgs e)
        {
            PreviewTimer.Stop();
            SetProgressingRingStatus(true);

            var source = (ImageItem)SearchListBox.SelectedValue;
            var targetStyle = PreviewListBox.SelectedItems;
            if (source == null || targetStyle == null || targetStyle.Count == 0) return;

            if (source.FullSizeImageFile != null)
            {
                ApplyStyle();
                SetProgressingRingStatus(false);
            }
            else if (!_applyDownloadingUriList.Contains(source.FullSizeImageUri))
            {
                var fullsizeImageUri = source.FullSizeImageUri;
                _customizeDownloadingUriList.Remove(fullsizeImageUri);
                _applyDownloadingUriList.Add(fullsizeImageUri);

                var fullsizeImageFile = TempPath.GetPath("fullsize");
                new Downloader()
                    .Get(fullsizeImageUri, fullsizeImageFile)
                    .After(() => { HandleDownloadedFullSizeImage(source, fullsizeImageFile); })
                    .OnError(() =>
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            var currentImageItem = SearchListBox.SelectedValue as ImageItem;
                            if (currentImageItem == null)
                            {
                                SetProgressingRingStatus(false);
                            }
                            else if (currentImageItem.ImageFile == source.ImageFile)
                            {
                                ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorNetworkOrSourceUnavailable);
                            }
                        }));
                    })
                    .Start();
            }
        }

        private void ApplyStyle()
        {
            if (PreviewListBox.SelectedValue == null) return;

            var source = SearchListBox.SelectedValue as ImageItem;
            Assumption.Made(source != null, "source item is null/empty");

            try
            {
                PreviewPresentation.ApplyStyle(source);
                this.ShowMessageAsync("", TextCollection.ImagesLabText.SuccessfullyAppliedStyle);
            }
            catch (AssumptionFailedException)
            {
                ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorNoSelectedSlide);
            }
        }

        private void StyleCustomizeButton_OnClick(object sender, RoutedEventArgs e)
        {
            PreviewTimer.Stop();
            SetProgressingRingStatus(true);

            var source = (ImageItem)SearchListBox.SelectedValue;
            var targetStyle = PreviewListBox.SelectedItems;
            if (source == null || targetStyle == null || targetStyle.Count == 0) return;

            if (source.FullSizeImageFile != null)
            {
                OpenCustomizationFlyout();
                SetProgressingRingStatus(false);
            }
            else if (!_customizeDownloadingUriList.Contains(source.FullSizeImageUri))
            {
                var fullsizeImageUri = source.FullSizeImageUri;
                _applyDownloadingUriList.Remove(fullsizeImageUri);
                _customizeDownloadingUriList.Add(fullsizeImageUri);

                var fullsizeImageFile = TempPath.GetPath("fullsize");
                new Downloader()
                    .Get(fullsizeImageUri, fullsizeImageFile)
                    .After(() => { HandleDownloadedFullSizeImage(source, fullsizeImageFile); })
                    .OnError(() =>
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            var currentImageItem = SearchListBox.SelectedValue as ImageItem;
                            if (currentImageItem == null)
                            {
                                SetProgressingRingStatus(false);
                            }
                            else if (currentImageItem.ImageFile == source.ImageFile)
                            {
                                ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorNetworkOrSourceUnavailable);
                            }
                        }));
                    })
                    .Start();
            }
        }

        private void OpenCustomizationFlyout()
        {
            var source = (ImageItem)SearchListBox.SelectedValue;
            var targetStyles = PreviewListBox.SelectedItems;
            if (source == null || targetStyles == null || targetStyles.Count == 0) return;
            OpenCustomizationFlyout(targetStyles);
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
                TimeSpan.FromMilliseconds(600))
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
            if (_isVariationsFlyoutOpen) return;

            var left2RightToShowTranslate = new TranslateTransform { X = -StylesPreviewGrid.ActualWidth };
            StyleVariationsFlyout.RenderTransform = left2RightToShowTranslate;
            StyleVariationsFlyout.Visibility = Visibility.Visible;
            var left2RightToShowAnimation = new DoubleAnimation(-StylesPreviewGrid.ActualWidth, 0,
                TimeSpan.FromMilliseconds(600))
            {
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut },
                AccelerationRatio = 0.5
            };

            left2RightToShowTranslate.BeginAnimation(TranslateTransform.XProperty, left2RightToShowAnimation);
            _isVariationsFlyoutOpen = true;
        }
    }
}
