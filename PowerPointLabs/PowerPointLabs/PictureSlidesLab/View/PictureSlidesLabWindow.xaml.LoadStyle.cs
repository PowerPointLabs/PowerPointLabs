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
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.PictureSlidesLab.View
{
    partial class PictureSlidesLabWindow
    {
        private readonly SlideSelectionDialog _loadFromSlideDialog = new SlideSelectionDialog();
        private readonly SlideSelectionDialog _loadUserCustomizedStyleDialog = new SlideSelectionDialog();

        private const string ShapeNamePrefix = EffectsDesigner.ShapeNamePrefix;

        private void LoadFromSlideButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (_loadFromSlideDialog.IsOpen) return;

            _loadFromSlideDialog
                .Init("Load Style or Picture from the Selected Slide")
                .CustomizeGotoSlideButton("Load Style", "Load style from the selected slide.")
                .CustomizeAdditionalButton("Load Picture", "Load picture from the selected slide.")
                .FocusOkButton()
                .OpenDialog();
            this.ShowMetroDialogAsync(_loadFromSlideDialog, MetroDialogOptions);
        }

        private void LoadUserCustomizedStyleButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (_loadUserCustomizedStyleDialog.IsOpen) return;

            if (ViewModel.UserCustomizedStyles.Count == 0)
            {
                ShowInfoMessageBox(TextCollection.PictureSlidesLabText.ErrorNoUserCustomizedStyleInfo);
            }
            else
            {
                _loadUserCustomizedStyleDialog
                    .Init(GenerateUserCustomizedStylesImageItem(), "Load User-customized Style")
                    .CustomizeGotoSlideButton("Load", "Load user-customized style.")
                    .CustomizeAdditionalButton("Delete", "Delete user-customized style.")
                    .FocusOkButton()
                    .OpenDialog();
                this.ShowMetroDialogAsync(_loadUserCustomizedStyleDialog, MetroDialogOptions);
            }
        }

        private List<ImageItem> GenerateUserCustomizedStylesImageItem()
        {
            var imageItems = new List<ImageItem>();

            var currentSlide = this.GetCurrentSlide();
            var originalShapeList = currentSlide.GetShapesWithPrefix(ShapeNamePrefix + "_" + EffectName.Original_DO_NOT_REMOVE);
            StyleOption userCustomizedStyle = null;

            if (originalShapeList.Count > 0)
            {
                var currentStyle = ConstructStyleFromShapeInfo(originalShapeList[0]);
                userCustomizedStyle = GetUserCustomizedStyle(currentStyle);
            }

            for (int i = 0; i < ViewModel.UserCustomizedStyles.Count; i++)
            {
                var currentUserCustomizedStyle = ViewModel.UserCustomizedStyles[i];

                var imageItem = ViewModel.GenerateImageItem(
                    (ImageItem)ImageSelectionListBox.SelectedValue ?? CreateDefaultPictureItem(),
                    currentSlide.GetNativeSlide(),
                    this.GetCurrentPresentation().SlideWidth,
                    this.GetCurrentPresentation().SlideHeight,
                    currentUserCustomizedStyle);

                if (userCustomizedStyle != null && currentUserCustomizedStyle == userCustomizedStyle)
                {
                    imageItem.Tooltip.Insert(0, "(Current) ");
                }

                imageItems.Add(imageItem);
            }

            return imageItems;
        }

        // it's actually using GotoSlide dialog, but to do stuff related to Load Styles
        private void InitLoadFromSlideDialog()
        {
            _loadFromSlideDialog.GetType()
                    .GetProperty("OwningWindow", BindingFlags.Instance | BindingFlags.NonPublic)
                    .SetValue(_loadFromSlideDialog, this, null);

            _loadFromSlideDialog.OnGotoSlide += LoadStyle;

            _loadFromSlideDialog.OnAdditionalButtonClick += LoadImage;

            _loadFromSlideDialog.OnCancel += () =>
            {
                _loadFromSlideDialog.CloseDialog();
                this.HideMetroDialogAsync(_loadFromSlideDialog, MetroDialogOptions);
            };
            Logger.Log("PSL init LoadFromSlideDialog done");
        }

        private void InitLoadUserCustomizedStyleDialog()
        {
            _loadUserCustomizedStyleDialog.GetType()
                    .GetProperty("OwningWindow", BindingFlags.Instance | BindingFlags.NonPublic)
                    .SetValue(_loadUserCustomizedStyleDialog, this, null);

            _loadUserCustomizedStyleDialog.OnGotoSlide += LoadUserCustomizedStyle;

            _loadUserCustomizedStyleDialog.OnAdditionalButtonClick += DeleteUserCustomizedStyle;

            _loadUserCustomizedStyleDialog.OnCancel += () =>
            {
                _loadUserCustomizedStyleDialog.CloseDialog();
                this.HideMetroDialogAsync(_loadUserCustomizedStyleDialog, MetroDialogOptions);
            };
            Logger.Log("PSL init LoadUserCustomizedStyleDialog done");
        }

        private void LoadImage()
        {
            _loadFromSlideDialog.CloseDialog();
            this.HideMetroDialogAsync(_loadFromSlideDialog, MetroDialogOptions);

            // which is the current slide
            var currentSlide = this.GetCurrentPresentation().Slides[_loadFromSlideDialog.SelectedSlide - 1];
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
                        UpdatePictureDimensionsInfo(croppedShapeList, originalImageShape, imageItem);
                        UpdateImageSelection(i);
                        break;
                    }
                }

                // if image source is deleted already, need to re-generate images
                // and put into listbox
                if (!isImageStillInListBox)
                {
                    var imageItem = ExtractImageItem(originalImageShape, croppedShapeList);
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
            _loadFromSlideDialog.CloseDialog();
            this.HideMetroDialogAsync(_loadFromSlideDialog, MetroDialogOptions);

            // which is the current slide
            var currentSlide = this.GetCurrentPresentation().Slides[_loadFromSlideDialog.SelectedSlide - 1];
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
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="targetSlide"></param>
        /// <param name="isLoadingWithDefaultPicture">when no style found, use default picture to preview style</param>
        /// <returns>is successfully loaded</returns>
        private bool LoadStyleAndImage(PowerPointSlide targetSlide, bool isLoadingWithDefaultPicture = true)
        {
            if (targetSlide == null) return false;

            var isSuccessfullyLoaded = false;
            var originalShapeList = targetSlide
                .GetShapesWithPrefix(ShapeNamePrefix + "_" + EffectName.Original_DO_NOT_REMOVE);
            var croppedShapeList = targetSlide
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
                    var imageItem = ExtractImageItem(originalImageShape, croppedShapeList);
                    ViewModel.ImageSelectionList.Add(imageItem);

                    ImageSelectionListBox.SelectedIndex = ImageSelectionListBox.Items.Count - 1;
                    OpenVariationFlyoutForReload(styleName, originalImageShape);
                }
                isSuccessfullyLoaded = true;
            }
            return isSuccessfullyLoaded;
        }

        private void LoadUserCustomizedStyle()
        {
            _loadUserCustomizedStyleDialog.CloseDialog();
            this.HideMetroDialogAsync(_loadUserCustomizedStyleDialog, MetroDialogOptions);

            if (ImageSelectionListBox.SelectedIndex < 0)
            {
                UpdatePreviewImages(CreateDefaultPictureItem());
            }
            else
            {
                UpdatePreviewImages((ImageItem)ImageSelectionListBox.SelectedValue);
            }

            var styleOption = ViewModel.UserCustomizedStyles[_loadUserCustomizedStyleDialog.SelectedSlide];
            OpenVariationFlyoutForReload(styleOption, canUseDefaultPicture: true);
        }

        private void DeleteUserCustomizedStyle()
        {
            ViewModel.UserCustomizedStyles.RemoveAt(_loadUserCustomizedStyleDialog.SelectedSlide);
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

        private void OpenVariationFlyoutForReload(StyleOption styleOption, bool canUseDefaultPicture = false)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                StylesPreviewListBox.SelectedIndex = MapStyleNameToStyleIndex(styleOption.StyleName);
                var listOfStyles = ConstructStylesFromStyleOption(styleOption);
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
                            (ImageItem)ImageSelectionListBox.SelectedValue,
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
            var fullsizeImageFile = ExtractPictureInfo(originalImageShape);
            var fullsizeThumbnailFile = ImageUtil.GetThumbnailFromFullSizeImg(fullsizeImageFile);
            // get dimensions/cropped picture info
            var croppedImageFile = ExtractCroppedPicture(croppedShapeList);
            var croppedThumbnailFile = ImageUtil.GetThumbnailFromFullSizeImg(croppedImageFile);
            var rect = ExtractDimensionsInfo(originalImageShape);

            // then form image item
            var imageItem = new ImageItem
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
            var fullsizeImageFile =
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
            var croppedImageFile = ExtractCroppedPicture(croppedShapeList);
            var croppedThumbnailFile = ImageUtil.GetThumbnailFromFullSizeImg(croppedImageFile);
            var rect = ExtractDimensionsInfo(originalImageShape);
            imageItem.CroppedImageFile = croppedImageFile;
            imageItem.CroppedThumbnailImageFile = croppedThumbnailFile;
            imageItem.Rect = rect;
        }

        private static string ExtractCroppedPicture(List<Shape> croppedShapeList)
        {
            var croppedImageFile =
                StoragePath.GetPath("crop-" + DateTime.Now.GetHashCode() +
                                    Guid.NewGuid().ToString().Substring(0, 7) + ".jpg");
            var croppedImageShape = croppedShapeList.Count > 0 ? croppedShapeList[0] : null;
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
            var rect = new Rect();
            rect.X = double.Parse(originalImageShape.Tags[Service.Effect.Tag.ReloadRectX]);
            rect.Y = double.Parse(originalImageShape.Tags[Service.Effect.Tag.ReloadRectY]);
            rect.Width = double.Parse(originalImageShape.Tags[Service.Effect.Tag.ReloadRectWidth]);
            rect.Height = double.Parse(originalImageShape.Tags[Service.Effect.Tag.ReloadRectHeight]);
            return rect;
        }

        private int MapStyleNameToStyleIndex(string styleName)
        {
            var allOptions = ViewModel.OptionsFactory.GetAllStylesPreviewOptions();
            for (var i = 0; i < allOptions.Count; i++)
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
            var result = new List<StyleOption>();
            for (var i = 0; i < 8; i++)
            {
                result.Add(ConstructStyleFromShapeInfo(shape));
            }
            return result;
        }

        private List<StyleOption> ConstructStylesFromStyleOption(StyleOption styleOption)
        {
            var result = new List<StyleOption>();
            for (var i = 0; i < 8; i++)
            {
                result.Add(ConstructStyleFromStyleOption(styleOption));
            }
            return result;
        }

        private Dictionary<string, List<StyleVariant>> ConstructVariantsFromStyle(StyleOption opt)
        {
            var variants = ViewModel.VariantsFactory.GetVariants(opt.StyleName);
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

        private StyleOption ConstructStyleFromShapeInfo(Shape shape)
        {
            var opt = new StyleOption();
            var props = opt.GetType().GetProperties();
            foreach (var propertyInfo in props)
            {
                var valueInStr = shape.Tags[Service.Effect.Tag.ReloadPrefix + propertyInfo.Name];
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

        private StyleOption ConstructStyleFromStyleOption(StyleOption styleOption)
        {
            var opt = new StyleOption();
            var props = opt.GetType().GetProperties();
            foreach (var propertyInfo in props)
            {
                var value = styleOption.GetType().GetProperty(propertyInfo.Name).GetValue(styleOption, null);
                propertyInfo.SetValue(opt, value, null);
            }
            return opt;
        }

        private StyleOption GetUserCustomizedStyle(StyleOption styleOption)
        {
            var type = styleOption.GetType();

            for (int i = 0; i < ViewModel.UserCustomizedStyles.Count; i++)
            {
                var style = ViewModel.UserCustomizedStyles[i];
                var styleType = style.GetType();
                var isFound = true;

                foreach (var propertyInfo in type.GetProperties())
                {
                    if (propertyInfo.Name != "OptionName"
                        && !propertyInfo.GetValue(styleOption, null).Equals(styleType.GetProperty(propertyInfo.Name).GetValue(style, null)))
                    {
                        isFound = false;
                        break;
                    }
                }

                if (isFound)
                {
                    return style;
                }
            }

            return null;
        }
    }
    #endregion
}
