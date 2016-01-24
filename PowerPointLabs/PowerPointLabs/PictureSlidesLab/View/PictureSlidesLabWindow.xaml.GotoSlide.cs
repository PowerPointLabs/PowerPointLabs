using System;
using System.Reflection;
using System.Windows;
using MahApps.Metro.Controls.Dialogs;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.PictureSlidesLab.Util;

namespace PowerPointLabs.PictureSlidesLab.View
{
    partial class PictureSlidesLabWindow
    {
        private readonly SlideSelectionDialog _gotoSlideDialog = new SlideSelectionDialog();
        private bool _isDisplayDefaultPicture;

        private void GotoSlideButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                _gotoSlideDialog.Init("Go to the Selected Slide");
                _gotoSlideDialog.CustomizeAdditionalButton("Go directly", 
                    "Go to the selected slide without changing the current style.");
                _gotoSlideDialog.FocusOkButton();
                this.ShowMetroDialogAsync(_gotoSlideDialog, MetroDialogOptions);
            }
            catch
            {
                // dialog could be fired multiple times
            }
        }

        private void InitGotoSlideDialog()
        {
            _gotoSlideDialog.GetType()
                    .GetProperty("OwningWindow", BindingFlags.Instance | BindingFlags.NonPublic)
                    .SetValue(_gotoSlideDialog, this, null);

            _gotoSlideDialog.OnGotoSlide += GotoSlideWithStyleLoading;

            _gotoSlideDialog.OnAdditionalButtonClick += GotoSlideDirectly;

            _gotoSlideDialog.OnCancel += () =>
            {
                this.HideMetroDialogAsync(_gotoSlideDialog, MetroDialogOptions);
            };
        }

        private void GotoSlideWithStyleLoading()
        {
            this.HideMetroDialogAsync(_gotoSlideDialog, MetroDialogOptions);

            GotoSlide();

            // which is the current slide
            var currentSlide = PowerPointPresentation.Current.Slides[_gotoSlideDialog.SelectedSlide - 1];
            if (currentSlide == null) return;

            var originalShapeList = currentSlide.GetShapesWithPrefix(ShapeNamePrefix + "_" + EffectName.Original_DO_NOT_REMOVE);
            var croppedShapeList = currentSlide.GetShapesWithPrefix(ShapeNamePrefix + "_" + EffectName.Cropped_DO_NOT_REMOVE);

            // if no original shape, show default picture
            if (originalShapeList.Count == 0)
            {
                _isDisplayDefaultPicture = true;
                ImageSelectionListBox.SelectedIndex = -1;
                _isDisplayDefaultPicture = false;
                UpdatePreviewImages(CreateDefaultPictureItem());
                _isDisplayDefaultPicture = true;
            }
            else // load the style
            {
                var originalImageShape = originalShapeList[0];
                var isImageStillInListBox = false;
                var styleName = originalImageShape.Tags[Service.Effect.Tag.ReloadPrefix + "StyleName"];

                // if the image source is still in the listbox,
                // select it as source and also select the target style
                for (var i = 0; i < ImageSelectionListBox.Items.Count; i++)
                {
                    var imageItem = (ImageItem)ImageSelectionListBox.Items[i];
                    if (imageItem.FullSizeImageFile == originalImageShape.Tags[Service.Effect.Tag.ReloadOriginImg]
                        || imageItem.ContextLink == originalImageShape.Tags[Service.Effect.Tag.ReloadImgContext])
                    {
                        isImageStillInListBox = true;
                        ImageSelectionListBox.SelectedIndex = i;
                        // previewing is done async, need to use beginInvoke
                        // so that it's after previewing
                        OpenVariationFlyoutForReload(styleName, originalImageShape);
                        break;
                    }
                }

                // if image source is deleted already, need to re-generate images
                // and put into listbox
                if (!isImageStillInListBox)
                {
                    var fullsizeImageFile =
                        StoragePath.GetPath("img-" + DateTime.Now.GetHashCode() +
                                            Guid.NewGuid().ToString().Substring(0, 7) + ".jpg");
                    // need to make shape visible so that can export
                    originalImageShape.Visible = MsoTriState.msoTrue;
                    originalImageShape.Export(fullsizeImageFile, PpShapeFormat.ppShapeFormatJPG);
                    originalImageShape.Visible = MsoTriState.msoFalse;

                    var fullsizeThumbnailFile = ImageUtil.GetThumbnailFromFullSizeImg(fullsizeImageFile);

                    var croppedImageFile =
                        StoragePath.GetPath("crop-" + DateTime.Now.GetHashCode() +
                                            Guid.NewGuid().ToString().Substring(0, 7) + ".jpg");
                    string croppedThumbnailFile;

                    var croppedImageShape = croppedShapeList.Count > 0 ? croppedShapeList[0] : null;
                    if (croppedImageShape != null)
                    {
                        croppedImageShape.Visible = MsoTriState.msoTrue;
                        croppedImageShape.Export(croppedImageFile, PpShapeFormat.ppShapeFormatJPG);
                        croppedThumbnailFile = ImageUtil.GetThumbnailFromFullSizeImg(croppedImageFile);
                        croppedImageShape.Visible = MsoTriState.msoFalse;
                    }
                    else
                    {
                        croppedImageFile = null;
                        croppedThumbnailFile = null;
                    }

                    var rect = new Rect();
                    rect.X = double.Parse(originalImageShape.Tags[Service.Effect.Tag.ReloadRectX]);
                    rect.Y = double.Parse(originalImageShape.Tags[Service.Effect.Tag.ReloadRectY]);
                    rect.Width = double.Parse(originalImageShape.Tags[Service.Effect.Tag.ReloadRectWidth]);
                    rect.Height = double.Parse(originalImageShape.Tags[Service.Effect.Tag.ReloadRectHeight]);

                    var imageItem = new ImageItem
                    {
                        ImageFile = fullsizeThumbnailFile,
                        FullSizeImageFile = fullsizeImageFile,
                        Tooltip = ImageUtil.GetWidthAndHeight(fullsizeImageFile),
                        CroppedImageFile = croppedImageFile,
                        CroppedThumbnailImageFile = croppedThumbnailFile,
                        ContextLink = originalImageShape.Tags[Service.Effect.Tag.ReloadImgContext],
                        Rect = rect
                    };

                    ViewModel.ImageSelectionList.Add(imageItem);

                    ImageSelectionListBox.SelectedIndex = ImageSelectionListBox.Items.Count - 1;
                    OpenVariationFlyoutForReload(styleName, originalImageShape);
                }
            }
        }

        private void GotoSlideDirectly()
        {
            this.HideMetroDialogAsync(_gotoSlideDialog, MetroDialogOptions);
            GotoSlide();
        }

        private void GotoSlide()
        {
            if (PowerPointCurrentPresentationInfo.CurrentSlide == null
                || _gotoSlideDialog.SelectedSlide != PowerPointCurrentPresentationInfo.CurrentSlide.Index)
            {
                PowerPointPresentation.Current.GotoSlide(_gotoSlideDialog.SelectedSlide);
            }
            UpdatePreviewImages();
        }

        private ImageItem CreateDefaultPictureItem()
        {
            return new ImageItem
            {
                ImageFile = StoragePath.NoPicturePlaceholderImgPath,
                Tooltip = "Please select a picture."
            };
        }
    }
}
