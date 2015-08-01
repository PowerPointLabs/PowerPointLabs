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
        private void DoPreview(Action beforePreview = null)
        {
            var image = (ImageItem)SearchListBox.SelectedValue;
            if (image == null || image.ImageFile == TempPath.LoadingImgPath)
            {
                PreviewList.Clear();
                PreviewProgressRing.IsActive = false;
            }
            else
            {
                if (beforePreview != null) beforePreview();

                PreviewTimer.Stop();
                DoPreview(image);
                // when timer ticks, try to download full size image to replace
                PreviewTimer.Start();
            }
        }

        private void DoPreview(ImageItem source)
        {
            // ui thread
            Dispatcher.BeginInvoke(new Action(() =>
            {
                var previousTextCopy = Clipboard.GetText();
                var selectedId = PreviewListBox.SelectedIndex;
                PreviewList.Clear();

                if (PowerPointCurrentPresentationInfo.CurrentSlide != null)
                {
                    var previewInfo = PreviewPresentation.PreviewStyles(source);
                    Add(PreviewList, previewInfo.DirectTextStyleImagePath, TextCollection.ImagesLabText.StyleNameDirectText);
                    Add(PreviewList, previewInfo.BlurStyleImagePath, TextCollection.ImagesLabText.StyleNameBlur);
                    Add(PreviewList, previewInfo.TextboxStyleImagePath, TextCollection.ImagesLabText.StyleNameTextBox);
                    Add(PreviewList, previewInfo.BannerStyleImagePath, TextCollection.ImagesLabText.StyleNameBanner);
                    Add(PreviewList, previewInfo.SpecialEffectStyleImagePath, TextCollection.ImagesLabText.StyleNameSpecialEffect);

                    PreviewListBox.SelectedIndex = selectedId;
                }
                if (previousTextCopy.Length > 0)
                {
                    Clipboard.SetText(previousTextCopy);
                }

                UpdateConfirmApplyPreviewImage();
                PreviewProgressRing.IsActive = false;
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
