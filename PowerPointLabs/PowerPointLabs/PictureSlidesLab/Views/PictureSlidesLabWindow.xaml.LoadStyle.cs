using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows;

using MahApps.Metro.Controls.Dialogs;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.PictureSlidesLab.Util;
using PowerPointLabs.TextCollection;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.PictureSlidesLab.Views
{
    partial class PictureSlidesLabWindow
    {
        private readonly SlideSelectionDialog _loadStylesDialog = new SlideSelectionDialog();

        private const string ShapeNamePrefix = EffectsDesigner.ShapeNamePrefix;

        private void LoadButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (_loadStylesDialog.IsOpen)
            {
                return;
            }

            _loadStylesDialog
                .Init("Load Style or Picture from the Selected Slide")
                .CustomizeGotoSlideButton("Load Style", "Load style from the selected slide.")
                .CustomizeAdditionalButton("Load Picture", "Load picture from the selected slide.")
                .FocusOkButton()
                .OpenDialog();
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
                _loadStylesDialog.CloseDialog();
                this.HideMetroDialogAsync(_loadStylesDialog, MetroDialogOptions);
            };
            Logger.Log("PSL init LoadStylesDialog done");
        }

        private void LoadImage()
        {
            _loadStylesDialog.CloseDialog();
            this.HideMetroDialogAsync(_loadStylesDialog, MetroDialogOptions);

            // which is the current slide
            PowerPointSlide currentSlide = this.GetCurrentPresentation().Slides[_loadStylesDialog.SelectedSlide - 1];
            if (currentSlide == null)
            {
                return;
            }

            List<Shape> originalShapeList = currentSlide.GetShapesWithPrefix(ShapeNamePrefix + "_" + EffectName.Original_DO_NOT_REMOVE);
            List<Shape> croppedShapeList = currentSlide.GetShapesWithPrefix(ShapeNamePrefix + "_" + EffectName.Cropped_DO_NOT_REMOVE);

            // if no original shape, show info
            if (originalShapeList.Count == 0)
            {
                ShowInfoMessageBox(PictureSlidesLabText.ErrorNoEmbeddedStyleInfo);
            }
            else
            {
                Shape originalImageShape = originalShapeList[0];
                bool isImageStillInListBox = false;

                // if the image source is still in the listbox,
                // select it as source and also select the target style
                for (int i = 0; i < ImageSelectionListBox.Items.Count; i++)
                {
                    ImageItem imageItem = (ImageItem)ImageSelectionListBox.Items[i];
                    if (imageItem.FullSizeImageFile == originalImageShape.Tags[Service.Effect.Tag.ReloadOriginImg]
                        || imageItem.ContextLink == originalImageShape.Tags[Service.Effect.Tag.ReloadImgContext])
                    {
                        isImageStillInListBox = true;
                        UpdatePictureDimensionsInfo(croppedShapeList, originalImageShape, imageItem);
                        UpdateImageSelection(i);
                        break;
                    }
                }

                // if image source is deleted already, need to re-generate images
                // and put into listbox
                if (!isImageStillInListBox)
                {
                    ImageItem imageItem = ExtractImageItem(originalImageShape, croppedShapeList);
                    ViewModel.ImageSelectionList.Add(imageItem);
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        ImageSelectionListBox.SelectedIndex =
                            ViewModel.ImageSelectionList.Count - 1;
                    }));
                }
            }
        }

        private void LoadStyle()
        {
            _loadStylesDialog.CloseDialog();
            this.HideMetroDialogAsync(_loadStylesDialog, MetroDialogOptions);

            // which is the current slide
            PowerPointSlide currentSlide = this.GetCurrentPresentation().Slides[_loadStylesDialog.SelectedSlide - 1];
            if (currentSlide == null)
            {
                return;
            }

            List<Shape> originalShapeList = currentSlide.GetShapesWithPrefix(ShapeNamePrefix + "_" + EffectName.Original_DO_NOT_REMOVE);

            // if no original shape, show info
            if (originalShapeList.Count == 0)
            {
                ShowInfoMessageBox(PictureSlidesLabText.ErrorNoEmbeddedStyleInfo);
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

                Shape originalImageShape = originalShapeList[0];
                string styleName = originalImageShape.Tags[Service.Effect.Tag.ReloadPrefix + "StyleName"];
                OpenVariationFlyoutForReload(styleName, originalImageShape, canUseDefaultPicture: true);
            }
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="targetSlide"></param>
        /// <param name="isLoadingWithDefaultPicture">when no style found, use default picture to preview style</param>
        /// <returns>is successfully loaded</returns>
        private bool LoadStyleAndImage(PowerPointSlide targetSlide, bool isLoadingWithDefaultPicture = true)
        {
            if (targetSlide == null)
            {
                return false;
            }

            bool isSuccessfullyLoaded = false;
            List<Shape> originalShapeList = targetSlide
                .GetShapesWithPrefix(ShapeNamePrefix + "_" + EffectName.Original_DO_NOT_REMOVE);
            List<Shape> croppedShapeList = targetSlide
                .GetShapesWithPrefix(ShapeNamePrefix + "_" + EffectName.Cropped_DO_NOT_REMOVE);

            // if no original shape, show default picture
            if (originalShapeList.Count == 0 && isLoadingWithDefaultPicture)
            {
                // De-select the picture
                EnterDefaultPictureMode();

                UpdatePreviewImages(isEnteringPictureVariation: true);
                UpdatePreviewStageControls();
                isSuccessfullyLoaded = true;
            }
            else if (originalShapeList.Count > 0) // load the style
            {
                Logger.Log("Original shapes found.");
                Shape originalImageShape = originalShapeList[0];
                bool isImageStillInListBox = false;
                string styleName = originalImageShape.Tags[Service.Effect.Tag.ReloadPrefix + "StyleName"];

                // if the image source is still in the listbox,
                // select it as source and also select the target style
                for (int i = 0; i < ImageSelectionListBox.Items.Count; i++)
                {
                    ImageItem imageItem = (ImageItem)ImageSelectionListBox.Items[i];
                    if (imageItem.FullSizeImageFile == originalImageShape.Tags[Service.Effect.Tag.ReloadOriginImg]
                        || imageItem.ContextLink == originalImageShape.Tags[Service.Effect.Tag.ReloadImgContext])
                    {
                        isImageStillInListBox = true;
                        UpdatePictureDimensionsInfo(croppedShapeList, originalImageShape, imageItem);
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
                    ImageItem imageItem = ExtractImageItem(originalImageShape, croppedShapeList);
                    ViewModel.ImageSelectionList.Add(imageItem);

                    ImageSelectionListBox.SelectedIndex = ImageSelectionListBox.Items.Count - 1;
                    OpenVariationFlyoutForReload(styleName, originalImageShape);
                }
                isSuccessfullyLoaded = true;
            }
            return isSuccessfullyLoaded;
        }

        #region Helper funcs

        private void UpdateImageSelection(int indexToSelect)
        {
            if (ImageSelectionListBox.SelectedIndex != indexToSelect)
            {
                ImageSelectionListBox.SelectedIndex = indexToSelect;
            }
            else // same selection, need to update preview images manually
            {
                UpdatePreviewImages();
            }
        }

        private void OpenVariationFlyoutForReload(string styleName, Shape originalImageShape,
            bool canUseDefaultPicture = false)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                StylesPreviewListBox.SelectedIndex = MapStyleNameToStyleIndex(styleName);
                List<StyleOption> listOfStyles = ConstructStylesFromShapeInfo(originalImageShape);
                Dictionary<string, List<StyleVariant>> variants = ConstructVariantsFromStyle(listOfStyles[0]);
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
            // get picture info
            string fullsizeImageFile = ExtractPictureInfo(originalImageShape);
            string fullsizeThumbnailFile = ImageUtil.GetThumbnailFromFullSizeImg(fullsizeImageFile);
            // get dimensions/cropped picture info
            string croppedImageFile = ExtractCroppedPicture(croppedShapeList);
            string croppedThumbnailFile = ImageUtil.GetThumbnailFromFullSizeImg(croppedImageFile);
            Rect rect = ExtractDimensionsInfo(originalImageShape);

            // then form image item
            ImageItem imageItem = new ImageItem
            {
                ImageFile = fullsizeThumbnailFile,
                FullSizeImageFile = fullsizeImageFile,
                Tooltip = ImageUtil.GetWidthAndHeight(fullsizeImageFile),
                CroppedImageFile = croppedImageFile,
                CroppedThumbnailImageFile = croppedThumbnailFile,
                ContextLink = originalImageShape.Tags[Service.Effect.Tag.ReloadImgContext],
                Source = originalImageShape.Tags[Service.Effect.Tag.ReloadImgSource],
                Rect = rect
            };
            return imageItem;
        }

        private static string ExtractPictureInfo(Shape originalImageShape)
        {
            string fullsizeImageFile =
                StoragePath.GetPath("img-" + DateTime.Now.GetHashCode() +
                                    Guid.NewGuid().ToString().Substring(0, 7) + ".jpg");
            // need to make shape visible so that can export
            originalImageShape.Visible = MsoTriState.msoTrue;
            originalImageShape.Export(fullsizeImageFile, PpShapeFormat.ppShapeFormatJPG);
            originalImageShape.Visible = MsoTriState.msoFalse;
            return fullsizeImageFile;
        }

        private static void UpdatePictureDimensionsInfo(List<Shape> croppedShapeList, Shape originalImageShape, 
            ImageItem imageItem)
        {
            string croppedImageFile = ExtractCroppedPicture(croppedShapeList);
            string croppedThumbnailFile = ImageUtil.GetThumbnailFromFullSizeImg(croppedImageFile);
            Rect rect = ExtractDimensionsInfo(originalImageShape);
            imageItem.CroppedImageFile = croppedImageFile;
            imageItem.CroppedThumbnailImageFile = croppedThumbnailFile;
            imageItem.Rect = rect;
        }

        private static string ExtractCroppedPicture(List<Shape> croppedShapeList)
        {
            string croppedImageFile =
                StoragePath.GetPath("crop-" + DateTime.Now.GetHashCode() +
                                    Guid.NewGuid().ToString().Substring(0, 7) + ".jpg");
            Shape croppedImageShape = croppedShapeList.Count > 0 ? croppedShapeList[0] : null;
            if (croppedImageShape != null)
            {
                croppedImageShape.Visible = MsoTriState.msoTrue;
                croppedImageShape.Export(croppedImageFile, PpShapeFormat.ppShapeFormatJPG);
                croppedImageShape.Visible = MsoTriState.msoFalse;
            }
            else
            {
                croppedImageFile = null;
            }
            return croppedImageFile;
        }

        private static Rect ExtractDimensionsInfo(Shape originalImageShape)
        {
            Rect rect = new Rect();
            rect.X = double.Parse(originalImageShape.Tags[Service.Effect.Tag.ReloadRectX]);
            rect.Y = double.Parse(originalImageShape.Tags[Service.Effect.Tag.ReloadRectY]);
            rect.Width = double.Parse(originalImageShape.Tags[Service.Effect.Tag.ReloadRectWidth]);
            rect.Height = double.Parse(originalImageShape.Tags[Service.Effect.Tag.ReloadRectHeight]);
            return rect;
        }

        private int MapStyleNameToStyleIndex(string styleName)
        {
            List<StyleOption> allOptions = ViewModel.OptionsFactory.GetAllStylesPreviewOptions();
            for (int i = 0; i < allOptions.Count; i++)
            {
                if (allOptions[i].StyleName == styleName)
                {
                    return i;
                }
            }
            return 0;
        }

        private List<StyleOption> ConstructStylesFromShapeInfo(Shape shape)
        {
            List<StyleOption> result = new List<StyleOption>();
            for (int i = 0; i < 8; i++)
            {
                result.Add(ConstructStyleFromShapeInfo(shape));
            }
            return result;
        }

        private Dictionary<string, List<StyleVariant>> ConstructVariantsFromStyle(StyleOption opt)
        {
            Dictionary<string, List<StyleVariant>> variants = ViewModel.VariantsFactory.GetVariants(opt.StyleName);
            // replace each category/aspect's variant
            // with the new variant from the given style options
            foreach (KeyValuePair<string, List<StyleVariant>> pair in variants)
            {
                StyleVariant firstVariant = pair.Value[0];
                StyleVariant newFirstVariant = firstVariant.Copy(opt);
                for (int i = 0; i < pair.Value.Count; i++)
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

        private StyleOption ConstructStyleFromShapeInfo(Shape shape)
        {
            StyleOption opt = new StyleOption();
            PropertyInfo[] props = opt.GetType().GetProperties();
            foreach (PropertyInfo propertyInfo in props)
            {
                string valueInStr = shape.Tags[Service.Effect.Tag.ReloadPrefix + propertyInfo.Name];
                if (propertyInfo.PropertyType == typeof(string))
                {
                    propertyInfo.SetValue(opt, valueInStr, null);
                }
                else if (propertyInfo.PropertyType == typeof(int))
                {
                    int valueInInt = int.Parse(valueInStr);
                    propertyInfo.SetValue(opt, valueInInt, null);
                }
                else if (propertyInfo.PropertyType == typeof(bool))
                {
                    bool valueInBool = bool.Parse(valueInStr);
                    propertyInfo.SetValue(opt, valueInBool, null);
                }
            }
            return opt;
        }
    }
    #endregion
}
