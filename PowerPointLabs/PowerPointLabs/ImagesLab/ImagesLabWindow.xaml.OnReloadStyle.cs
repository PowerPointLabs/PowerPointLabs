using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows;
using MahApps.Metro.Controls.Dialogs;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ImagesLab.Domain;
using PowerPointLabs.ImagesLab.Handler.Effect;
using PowerPointLabs.ImagesLab.Util;
using PowerPointLabs.Models;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.ImagesLab
{
    partial class ImagesLabWindow
    {
        private readonly SlideSelectionDialog _reloadStylesDialog = new SlideSelectionDialog();

        private const string ShapeNamePrefix = "pptImagesLab";

        private void ReloadButton_OnClick(object sender, RoutedEventArgs e)
        {
            // TODO move this to text collection
            _reloadStylesDialog.Init("Load Styles or Image from the Selected Slide");
            _reloadStylesDialog.CustomizeGotoSlideButton("Load Styles", "Load styles and image from the selected slide.");
            _reloadStylesDialog.CustomizeAdditionalButton("Load Image", "Load image from the selected slide.");
            _reloadStylesDialog.FocusOkButton();
            this.ShowMetroDialogAsync(_reloadStylesDialog, MetroDialogOptions);
        }

        // actually using GotoSlide dialog, but to do stuff related to Reload Styles
        private void InitReloadStylesDialog()
        {
            _reloadStylesDialog.GetType()
                    .GetProperty("OwningWindow", BindingFlags.Instance | BindingFlags.NonPublic)
                    .SetValue(_reloadStylesDialog, this, null);

            _reloadStylesDialog.OnGotoSlide += ReloadStyles();

            _reloadStylesDialog.OnAdditionalButtonClick += ReloadStyles(isReloadImageOnly: true);

            _reloadStylesDialog.OnCancel += () =>
            {
                this.HideMetroDialogAsync(_reloadStylesDialog, MetroDialogOptions);
            };
        }

        private SlideSelectionDialog.OkEvent ReloadStyles(bool isReloadImageOnly = false)
        {
            return () =>
            {
                this.HideMetroDialogAsync(_reloadStylesDialog, MetroDialogOptions);
                // go to the target slide
                if (_reloadStylesDialog.SelectedSlide > 0)
                {
                    PowerPointPresentation.Current.GotoSlide(_reloadStylesDialog.SelectedSlide);
                }

                // which is the current slide
                var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
                if (currentSlide == null) return;

                var originalShapeList = currentSlide.GetShapesWithPrefix(ShapeNamePrefix + "_" + EffectName.Original_DO_NOT_REMOVE);
                var croppedShapeList = currentSlide.GetShapesWithPrefix(ShapeNamePrefix + "_" + EffectName.Cropped_DO_NOT_REMOVE);

                // if no original shape, show info
                if (originalShapeList.Count == 0)
                {
                    ShowInfoMessageBox("No Images Lab styles are detected for the current slide.");
                }
                else
                {
                    var originalImageShape = originalShapeList[0];
                    var isImageStillInListBox = false;
                    var styleName = originalImageShape.Tags[Handler.Effect.Tag.ReloadPrefix + "StyleName"];

                    // if the image source is still in the listbox,
                    // select it as source and also select the target style
                    for(var i = 0; i < ImageSelectionListBox.Items.Count; i++)
                    {
                        var imageItem = (ImageItem) ImageSelectionListBox.Items[i];
                        if (imageItem.FullSizeImageFile
                            == originalImageShape.Tags[Handler.Effect.Tag.ReloadOriginImg])
                        {
                            isImageStillInListBox = true;
                            Dispatcher.Invoke(new Action(() =>
                            {
                                ImageSelectionListBox.SelectedIndex = i;
                                // previewing is done async, need to use beginInvoke
                                // so that it's after previewing
                                if (!isReloadImageOnly)
                                {
                                    OpenVariationFlyoutForReload(styleName, originalImageShape);
                                }
                            }));
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
                        rect.X = double.Parse(originalImageShape.Tags[Handler.Effect.Tag.ReloadRectX]);
                        rect.Y = double.Parse(originalImageShape.Tags[Handler.Effect.Tag.ReloadRectY]);
                        rect.Width = double.Parse(originalImageShape.Tags[Handler.Effect.Tag.ReloadRectWidth]);
                        rect.Height = double.Parse(originalImageShape.Tags[Handler.Effect.Tag.ReloadRectHeight]);

                        var imageItem = new ImageItem
                        {
                            ImageFile = fullsizeThumbnailFile,
                            FullSizeImageFile = fullsizeImageFile,
                            FullSizeImageUri = fullsizeImageFile,
                            Tooltip = ImageUtil.GetWidthAndHeight(fullsizeImageFile),
                            CroppedImageFile = croppedImageFile,
                            CroppedThumbnailImageFile = croppedThumbnailFile,
                            Rect = rect
                        };

                        Dispatcher.Invoke(new Action(() =>
                        {
                            ImageSelectionList.Add(imageItem);

                            if (!isReloadImageOnly)
                            {
                                ImageSelectionListBox.SelectedIndex = ImageSelectionListBox.Items.Count - 1;
                                OpenVariationFlyoutForReload(styleName, originalImageShape);
                            }
                        }));
                    }
                }
            };
        }

        private void OpenVariationFlyoutForReload(string styleName, Shape originalImageShape)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                StylesPreviewListBox.SelectedIndex = MapStyleNameToStyleIndex(styleName);
                var listOfStyles = ConstructStylesFromShapeInfo(originalImageShape);
                var variants = ConstructVariantsFromStyle(listOfStyles[0]);
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    CustomizeStyle(listOfStyles, variants);
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        StylesVariationListBox.ScrollIntoView(StylesVariationListBox.SelectedItem);
                    }));
                }));
            }));
        }

        private int MapStyleNameToStyleIndex(string styleName)
        {
            switch (styleName)
            {
                case TextCollection.ImagesLabText.StyleNameDirectText:
                    return TextCollection.ImagesLabText.StyleIndexDirectText;
                case TextCollection.ImagesLabText.StyleNameBlur:
                    return TextCollection.ImagesLabText.StyleIndexBlur;
                case TextCollection.ImagesLabText.StyleNameTextBox:
                    return TextCollection.ImagesLabText.StyleIndexTextBox;
                case TextCollection.ImagesLabText.StyleNameBanner:
                    return TextCollection.ImagesLabText.StyleIndexBanner;
                case TextCollection.ImagesLabText.StyleNameSpecialEffect:
                    return TextCollection.ImagesLabText.StyleIndexSpecialEffect;
                case TextCollection.ImagesLabText.StyleNameOverlay:
                    return TextCollection.ImagesLabText.StyleIndexOverlay;
                default:
                    return 0;
            }
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
                var valueInStr = shape.Tags[Handler.Effect.Tag.ReloadPrefix + propertyInfo.Name];
                if (string.IsNullOrEmpty(valueInStr))
                {
                    continue;
                }

                if (propertyInfo.PropertyType == typeof (string))
                {
                    propertyInfo.SetValue(opt, valueInStr, null);
                }
                else if (propertyInfo.PropertyType == typeof (int))
                {
                    var valueInInt = int.Parse(valueInStr);
                    propertyInfo.SetValue(opt, valueInInt, null);
                }
                else if (propertyInfo.PropertyType == typeof (bool))
                {
                    var valueInBool = bool.Parse(valueInStr);
                    propertyInfo.SetValue(opt, valueInBool, null);
                }
            }
            return opt;
        }
    }
}
