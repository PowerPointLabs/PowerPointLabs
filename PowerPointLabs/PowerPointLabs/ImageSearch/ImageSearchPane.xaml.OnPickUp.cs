using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using MahApps.Metro.Controls.Dialogs;
using PowerPointLabs.AutoUpdate;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Exceptions;

namespace PowerPointLabs.ImageSearch
{
    partial class ImageSearchPane
    {
        private bool _isVariationsFlyoutOpen;

        private void UpdateStyleVariationsImages()
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

                    var styleOptions = StyleOptionsFactory.GetVariationOptions(targetStyle.Tooltip);
                    foreach (var styleOption in styleOptions)
                    {
                        UpdateStyleVariationsImage(styleOption, source);
                    }

                    VariationListBox.SelectedIndex = selectedId;
                    if (source.FullSizeImageFile != null)
                    {
                        SetProgressingRingStatus(false);
                    }
                }
                catch
                {
                    // ignore, selected slide may be null
                }
            }));
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

        private void PickUpStyle()
        {
            var source = (ImageItem)SearchListBox.SelectedValue;
            var targetStyle = PreviewListBox.SelectedItems;
            if (source == null || targetStyle == null || targetStyle.Count == 0) return;

            UpdateStyleVariationsImages();
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

                var targetStyle = ((ImageItem) PreviewListBox.SelectedValue).Tooltip;
                var options = StyleOptionsFactory.GetVariationOptions(targetStyle);

                var targetStyleOption = options[VariationListBox.SelectedIndex];
                targetStyleOption.PropertyChanged += (o, args) =>
                {
                    _latestStyleOptionsUpdateTime = DateTime.Now;
                };

                PreviewPresentation.SetStyleOptions(targetStyleOption);
                OptionsPane2.DataContext = targetStyleOption;
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
