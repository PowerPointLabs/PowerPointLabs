using System;
using System.Collections.Generic;
using System.Windows;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Models;

namespace PowerPointLabs.ImageSearch
{
    public partial class ImageSearchPane
    {
        private void DoPreview()
        {
            var image = (ImageItem)SearchListBox.SelectedValue;
            if (image == null || image.ImageFile == StoragePath.LoadingImgPath)
            {
                if (_isVariationsFlyoutOpen)
                {
                    VariationList.Clear();
                }
                else
                {
                    PreviewList.Clear();
                }
                SetProgressingRingStatus(false);
            }
            else if (_isVariationsFlyoutOpen)
            {
                PreviewTimer.Stop();
                UpdateStyleVariationsImages();
                PreviewTimer.Start();
            }
            else
            {
                PreviewTimer.Stop();
                DoPreview(image);
                _latestPreviewUpdateTime = DateTime.Now;
                // when timer ticks, try to download full size image to replace
                PreviewTimer.Start();
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
                    var selectedId = PreviewListBox.SelectedIndex;
                    PreviewList.Clear();

                    if (PowerPointCurrentPresentationInfo.CurrentSlide != null)
                    {
                        var previewInfo = PreviewPresentation.PreviewStyles(source);
                        Add(PreviewList, previewInfo.DirectTextStyleImagePath,
                            TextCollection.ImagesLabText.StyleNameDirectText);
                        Add(PreviewList, previewInfo.BlurStyleImagePath, TextCollection.ImagesLabText.StyleNameBlur);
                        Add(PreviewList, previewInfo.TextboxStyleImagePath,
                            TextCollection.ImagesLabText.StyleNameTextBox);
                        Add(PreviewList, previewInfo.BannerStyleImagePath, TextCollection.ImagesLabText.StyleNameBanner);
                        Add(PreviewList, previewInfo.SpecialEffectStyleImagePath,
                            TextCollection.ImagesLabText.StyleNameSpecialEffect);
                        Add(PreviewList, previewInfo.OverlayStyleImagePath,
                            TextCollection.ImagesLabText.StyleNameOverlay);

                        PreviewListBox.SelectedIndex = selectedId;
                        _latestPreviewUpdateTime = DateTime.Now;
                    }
                    if (previousTextCopy.Length > 0)
                    {
                        Clipboard.SetText(previousTextCopy);
                    }

                    if (_isCustomizationFlyoutOpen)
                    {
                        UpdateConfirmApplyPreviewImage();
                    }
                    else if (_isVariationsFlyoutOpen)
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
