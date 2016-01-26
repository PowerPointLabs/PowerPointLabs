using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows;
using MahApps.Metro.Controls.Dialogs;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory;
using PowerPointLabs.PictureSlidesLab.Service;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.PictureSlidesLab.Util;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.PictureSlidesLab.View
{
    partial class PictureSlidesLabWindow
    {
        private readonly SlideSelectionDialog _loadStylesDialog = new SlideSelectionDialog();

        private const string ShapeNamePrefix = EffectsDesigner.ShapeNamePrefix;

        private void LoadButton_OnClick(object sender, RoutedEventArgs e)
        {
            // TODO move this to text collection
            _loadStylesDialog.Init("Load Style or Image from the Selected Slide");
            _loadStylesDialog.CustomizeGotoSlideButton("Load Style", "Load style from the selected slide.");
            _loadStylesDialog.CustomizeAdditionalButton("Load Image", "Load image from the selected slide.");
            _loadStylesDialog.FocusOkButton();
            this.ShowMetroDialogAsync(_loadStylesDialog, MetroDialogOptions);
        }

        // it's actually using GotoSlide dialog, but to do stuff related to Load Styles
        private void InitLoadStylesDialog()
        {
            _loadStylesDialog.GetType()
                    .GetProperty("OwningWindow", BindingFlags.Instance | BindingFlags.NonPublic)
                    .SetValue(_loadStylesDialog, this, null);

            _loadStylesDialog.OnGotoSlide += LoadStyle;

            _loadStylesDialog.OnAdditionalButtonClick += LoadImage;

            _loadStylesDialog.OnCancel += () =>
            {
                this.HideMetroDialogAsync(_loadStylesDialog, MetroDialogOptions);
            };
        }

        private void LoadImage()
        {
            this.HideMetroDialogAsync(_loadStylesDialog, MetroDialogOptions);

            // which is the current slide
            var currentSlide = PowerPointPresentation.Current.Slides[_loadStylesDialog.SelectedSlide - 1];
            if (currentSlide == null) return;

            var originalShapeList = currentSlide.GetShapesWithPrefix(ShapeNamePrefix + "_" + EffectName.Original_DO_NOT_REMOVE);
            var croppedShapeList = currentSlide.GetShapesWithPrefix(ShapeNamePrefix + "_" + EffectName.Cropped_DO_NOT_REMOVE);

            // if no original shape, show info
            if (originalShapeList.Count == 0)
            {
                ShowInfoMessageBox(TextCollection.PictureSlidesLabText.ErrorNoEmbeddedStyleInfo);
            }
            else
            {
                var originalImageShape = originalShapeList[0];
                var isImageStillInListBox = false;

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
                        ShowInfoMessageBox(TextCollection.PictureSlidesLabText.SuccessfullyLoadedImage);
                        break;
                    }
                }

                // if image source is deleted already, need to re-generate images
                // and put into listbox
                if (!isImageStillInListBox)
                {
                    var imageItem = ExtractImageItem(originalImageShape, croppedShapeList);
                    ViewModel.ImageSelectionList.Add(imageItem);

                    ShowInfoMessageBox(TextCollection.PictureSlidesLabText.SuccessfullyLoadedImage);
                }
            }
        }

        private void LoadStyle()
        {
            this.HideMetroDialogAsync(_loadStylesDialog, MetroDialogOptions);

            // which is the current slide
            var currentSlide = PowerPointPresentation.Current.Slides[_loadStylesDialog.SelectedSlide - 1];
            if (currentSlide == null) return;

            var originalShapeList = currentSlide.GetShapesWithPrefix(ShapeNamePrefix + "_" + EffectName.Original_DO_NOT_REMOVE);

            // if no original shape, show info
            if (originalShapeList.Count == 0)
            {
                ShowInfoMessageBox(TextCollection.PictureSlidesLabText.ErrorNoEmbeddedStyleInfo);
            }
            else
            {
                if (ImageSelectionListBox.SelectedIndex < 0)
                {
                    UpdatePreviewImages(CreateDefaultPictureItem());
                }
                else
                {
                    UpdatePreviewImages((ImageItem) ImageSelectionListBox.SelectedValue);
                }

                var originalImageShape = originalShapeList[0];
                var styleName = originalImageShape.Tags[Service.Effect.Tag.ReloadPrefix + "StyleName"];
                OpenVariationFlyoutForReload(styleName, originalImageShape, canUseDefaultPicture: true);
            }
        }

        private void OpenVariationFlyoutForReload(string styleName, Shape originalImageShape,
            bool canUseDefaultPicture = false)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                StylesPreviewListBox.SelectedIndex = MapStyleNameToStyleIndex(styleName);
                var listOfStyles = ConstructStylesFromShapeInfo(originalImageShape);
                var variants = ConstructVariantsFromStyle(listOfStyles[0]);
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (canUseDefaultPicture
                        && ImageSelectionListBox.SelectedIndex < 0)
                    {
                        CustomizeStyle(CreateDefaultPictureItem(),
                            listOfStyles, variants);
                        EnterDefaultPictureMode();
                    }
                    else
                    {
                        CustomizeStyle(
                            (ImageItem) ImageSelectionListBox.SelectedValue,
                            listOfStyles, variants);
                    }
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        StylesVariationListBox.ScrollIntoView(StylesVariationListBox.SelectedItem);
                    }));
                }));
            }));
        }

        private static ImageItem ExtractImageItem(Shape originalImageShape, List<Shape> croppedShapeList)
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
            return imageItem;
        }

        private int MapStyleNameToStyleIndex(string styleName)
        {
            var allOptions = StyleOptionsFactory.GetAllStylesPreviewOptions();
            for (var i = 0; i < allOptions.Count; i++)
            {
                if (allOptions[i].StyleName == styleName)
                {
                    return i;
                }
            }
            return 0;
        }

        private List<StyleOptions> ConstructStylesFromShapeInfo(Shape shape)
        {
            var result = new List<StyleOptions>();
            for (var i = 0; i < 8; i++)
            {
                result.Add(ConstructStyleFromShapeInfo(shape));
            }
            return result;
        }

        private Dictionary<string, List<StyleVariants>> ConstructVariantsFromStyle(StyleOptions opt)
        {
            var variants = StyleVariantsFactory.GetVariants(opt.StyleName);
            // replace each category/aspect's variant
            // with the new variant from the given style options
            foreach (var pair in variants)
            {
                var firstVariant = pair.Value[0];
                var newFirstVariant = firstVariant.Copy(opt);
                for (var i = 0; i < pair.Value.Count; i++)
                {
                    // try to swap out the 'no-effect' style options
                    if (pair.Value[i].IsNoEffect(opt))
                    {
                        pair.Value[i] = firstVariant;
                        break;
                    }
                }
                pair.Value[0] = newFirstVariant;
            }
            return variants;
        }

        private StyleOptions ConstructStyleFromShapeInfo(Shape shape)
        {
            var opt = new StyleOptions();
            var props = opt.GetType().GetProperties();
            foreach (var propertyInfo in props)
            {
                var valueInStr = shape.Tags[Service.Effect.Tag.ReloadPrefix + propertyInfo.Name];
                if (string.IsNullOrEmpty(valueInStr))
                {
                    continue;
                }

                if (propertyInfo.PropertyType == typeof(string))
                {
                    propertyInfo.SetValue(opt, valueInStr, null);
                }
                else if (propertyInfo.PropertyType == typeof(int))
                {
                    var valueInInt = int.Parse(valueInStr);
                    propertyInfo.SetValue(opt, valueInInt, null);
                }
                else if (propertyInfo.PropertyType == typeof(bool))
                {
                    var valueInBool = bool.Parse(valueInStr);
                    propertyInfo.SetValue(opt, valueInBool, null);
                }
            }
            return opt;
        }
    }
}
