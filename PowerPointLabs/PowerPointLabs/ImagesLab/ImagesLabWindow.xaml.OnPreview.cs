using System;
using System.Collections.Generic;
using System.Windows;
using PowerPointLabs.ImagesLab.Domain;
using PowerPointLabs.ImagesLab.Factory;
using PowerPointLabs.ImagesLab.Util;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Exceptions;

namespace PowerPointLabs.ImagesLab
{
    public partial class ImagesLabWindow
    {
        private void DoPreview()
        {
            var image = (ImageItem)ImageSelectionListBox.SelectedValue;
            if (image == null || image.ImageFile == StoragePath.LoadingImgPath)
            {
                if (_isVariationsFlyoutOpen)
                {
                    StylesVariationList.Clear();
                }
                else
                {
                    StylesPreviewList.Clear();
                }
                SetProgressingRingStatus(false);
            }
            else if (_isVariationsFlyoutOpen)
            {
                UpdateStyleVariationsImages();
            }
            else
            {
                DoPreview(image);
                _latestPreviewUpdateTime = DateTime.Now;
            }
        }

        private void DoPreview(ImageItem source)
        {
            // ui thread
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    var previousTextCopy = Clipboard.GetText();
                    var selectedId = StylesPreviewListBox.SelectedIndex;
                    StylesPreviewList.Clear();

                    if (PowerPointCurrentPresentationInfo.CurrentSlide != null)
                    {
                        var allStylesPreviewOptions = StyleOptionsFactory.GetAllStylesPreviewOptions();
                        foreach (var stylesPreviewOption in allStylesPreviewOptions)
                        {
                            PreviewPresentation.SetStyleOptions(stylesPreviewOption);
                            var previewInfo = PreviewPresentation.PreviewApplyStyle(source);
                            Add(StylesPreviewList, previewInfo.PreviewApplyStyleImagePath,
                                stylesPreviewOption.StyleName);
                        }

                        StylesPreviewListBox.SelectedIndex = selectedId;
                        _latestPreviewUpdateTime = DateTime.Now;
                    }
                    if (previousTextCopy.Length > 0)
                    {
                        Clipboard.SetText(previousTextCopy);
                    }

                    if (_isVariationsFlyoutOpen)
                    {
                        UpdateStyleVariationsImages();
                    }
                    if (source.FullSizeImageFile != null)
                    {
                        SetProgressingRingStatus(false);
                    }
                }
                catch
                {
                    ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorImageCorrupted);
                    SetProgressingRingStatus(false);
                }
            }));
        }

        private void StylesPickUpButton_OnClick(object sender, RoutedEventArgs e)
        {
            CustomizeStyle();
        }

        private void StylesApplyButton_OnClick(object sender, RoutedEventArgs e)
        {
            ApplyStyleInPreviewStage();
        }

        private void ApplyStyleInPreviewStage()
        {
            if (StylesPreviewListBox.SelectedValue == null) return;

            var source = ImageSelectionListBox.SelectedValue as ImageItem;
            Assumption.Made(source != null, "source item is null/empty");

            try
            {
                var targetStyle = ((ImageItem) StylesPreviewListBox.SelectedValue).Tooltip;
                var targetDefaultOptions = StyleOptionsFactory.GetStylesPreviewOption(targetStyle);
                PreviewPresentation.SetStyleOptions(targetDefaultOptions);
                PreviewPresentation.ApplyStyle(source);

                OpenSuccessfullyAppliedDialog();
            }
            catch (AssumptionFailedException)
            {
                ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorNoSelectedSlide);
            }
        }

        private void Add(ICollection<ImageItem> list, string imagePath, string tooltip)
        {
            list.Add(new ImageItem
            {
                ImageFile = imagePath,
                Tooltip = tooltip
            });
        }
    }
}
