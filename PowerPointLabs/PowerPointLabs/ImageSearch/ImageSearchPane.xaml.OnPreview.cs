using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Models;

namespace PowerPointLabs.ImageSearch
{
    public partial class ImageSearchPane
    {
        private void DoPreview(IList<int> selectedIds = null)
        {
            var image = (ImageItem)SearchListBox.SelectedValue;
            if (image == null || image.ImageFile == TempPath.LoadingImgPath)
            {
                PreviewList.Clear();
                PreviewProgressRing.IsActive = false;
            }
            else
            {
                PreviewTimer.Stop();
                DoPreview(image, selectedIds);
                // when timer ticks, try to download full size image to replace
                PreviewTimer.Start();
            }
        }

        private void DoPreview(ImageItem source, IList<int> selectedIds = null)
        {
            // ui thread
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    var previousTextCopy = Clipboard.GetText();
                    selectedIds = selectedIds ?? GetSelectedIndices(PreviewListBox.SelectedItems);
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

                        SelectPreviewListBoxItems(selectedIds);
                        _latestPreviewUpdateTime = DateTime.Now;
                    }
                    if (previousTextCopy.Length > 0)
                    {
                        Clipboard.SetText(previousTextCopy);
                    }

                    UpdateConfirmApplyPreviewImage();
                    PreviewProgressRing.IsActive = false;
                }
                catch
                {
                    ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorImageCorrupted);
                    PreviewProgressRing.IsActive = false;
                }
            }));
        }

        // make it still select the same target styles after preview
        private void SelectPreviewListBoxItems(IList<int> selectedIds)
        {
            SelectPreviewListBox(
                TextCollection.ImagesLabText.StyleIndexDirectText,
                selectedIds.Any(val => val == TextCollection.ImagesLabText.StyleIndexDirectText));
            SelectPreviewListBox(
                TextCollection.ImagesLabText.StyleIndexBlur,
                selectedIds.Any(val => val == TextCollection.ImagesLabText.StyleIndexBlur));
            SelectPreviewListBox(
                TextCollection.ImagesLabText.StyleIndexTextBox,
                selectedIds.Any(val => val == TextCollection.ImagesLabText.StyleIndexTextBox));
            SelectPreviewListBox(
                TextCollection.ImagesLabText.StyleIndexBanner,
                selectedIds.Any(val => val == TextCollection.ImagesLabText.StyleIndexBanner));
            SelectPreviewListBox(
                TextCollection.ImagesLabText.StyleIndexSpecialEffect,
                selectedIds.Any(val => val == TextCollection.ImagesLabText.StyleIndexSpecialEffect));
        }

        // TODO extract this to somewhere COMMON
        private IList<int> GetSelectedIndices(IList items)
        {
            var result = new List<int>();
            foreach (ImageItem imageItem in items)
            {
                switch (imageItem.Tooltip)
                {
                    case TextCollection.ImagesLabText.StyleNameDirectText:
                        result.Add(TextCollection.ImagesLabText.StyleIndexDirectText);
                        break;
                    case TextCollection.ImagesLabText.StyleNameBlur:
                        result.Add(TextCollection.ImagesLabText.StyleIndexBlur);
                        break;
                    case TextCollection.ImagesLabText.StyleNameTextBox:
                        result.Add(TextCollection.ImagesLabText.StyleIndexTextBox);
                        break;
                    case TextCollection.ImagesLabText.StyleNameBanner:
                        result.Add(TextCollection.ImagesLabText.StyleIndexBanner);
                        break;
                    case TextCollection.ImagesLabText.StyleNameSpecialEffect:
                        result.Add(TextCollection.ImagesLabText.StyleIndexSpecialEffect);
                        break;
                }
            }
            return result;
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
