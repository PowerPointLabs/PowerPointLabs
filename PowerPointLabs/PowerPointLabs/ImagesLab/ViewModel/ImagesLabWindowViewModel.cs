using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using PowerPointLabs.AutoUpdate.Interface;
using PowerPointLabs.ImagesLab.Model;
using PowerPointLabs.ImagesLab.ModelFactory;
using PowerPointLabs.ImagesLab.Service;
using PowerPointLabs.ImagesLab.Thread;
using PowerPointLabs.ImagesLab.Util;
using PowerPointLabs.ImagesLab.View.Interface;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Exceptions;

namespace PowerPointLabs.ImagesLab.ViewModel
{
    class ImagesLabWindowViewModel
    {
        // UI model - list that holds image items
        public ObservableCollection<ImageItem> ImageSelectionList { get; set; }

        // UI model - list that holds styles preview items
        public ObservableCollection<ImageItem> StylesPreviewList { get; set; }

        // UI model - list that holds styles variations items
        public ObservableCollection<ImageItem> StylesVariationList { get; set; }

        // UI controller
        public IImagesLabWindow ImagesLabWindow { get; set; }

        // Downloader
        public IDownloader ImageDownloader { get; set; }

        // a background presentation that will do the preview processing
        // TODO: rename it
        private StylesDesigner PreviewPresentation { get; set; }

        // used to clean up unused image files
        private HashSet<string> ImageFilesInUse { get; set; }

        // for variation stage
        private string _previousVariantsCategory;
        private IList<StyleOptions> _styleOptions;
        private Dictionary<string, List<StyleVariants>> _styleVariants;

        public ImagesLabWindowViewModel(IImagesLabWindow imagesLabWindow)
        {
            ImagesLabWindow = imagesLabWindow;
            ImageDownloader = new ContextDownloader(ImagesLabWindow.GetThreadContext());

            StylesVariationList = new ObservableCollection<ImageItem>();
            StylesPreviewList = new ObservableCollection<ImageItem>();
            ImageSelectionList = StoragePath.Load();

            ImageFilesInUse = new HashSet<string>();
            foreach (var imageItem in ImageSelectionList)
            {
                ImageFilesInUse.Add(imageItem.ImageFile);
                ImageFilesInUse.Add(imageItem.FullSizeImageFile);
                if (imageItem.CroppedImageFile != null)
                {
                    ImageFilesInUse.Add(imageItem.CroppedImageFile);
                    ImageFilesInUse.Add(imageItem.CroppedThumbnailImageFile);
                }
            }

            var isTempPathInit = TempPath.InitTempFolder();
            var isStoragePathInit = StoragePath.InitPersistentFolder(ImageFilesInUse);
            if (!isTempPathInit || !isStoragePathInit)
            {
                ImagesLabWindow.ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorFailToInitTempFolder);
            }

            PreviewPresentation = new StylesDesigner();
            PreviewPresentation.Open(withWindow: false, focus: false);
        }

        public void CleanUp()
        {
            if (PreviewPresentation != null)
            {
                PreviewPresentation.Close();
            }
            StoragePath.Save(ImageSelectionList);
        }

        public void RemoveImageSelectionListItem(int index)
        {
            ImageSelectionList.RemoveAt(index);
        }

        public void ClearImageSelectionList()
        {
            ImageSelectionList.Clear();
        }

        public void ClearStyleVariationList()
        {
            StylesVariationList.Clear();
        }

        public void ClearStylesPreviewList()
        {
            StylesPreviewList.Clear();
        }

        public void AddImageSelectionListItem(ImageItem item)
        {
            ImageSelectionList.Add(item);
        }

        public void AddImageSelectionListItem(string[] filenames)
        {
            try
            {
                foreach (var filename in filenames)
                {
                    VerifyIsProperImage(filename);
                    var fromFileItem = new ImageItem
                    {
                        ImageFile = ImageUtil.GetThumbnailFromFullSizeImg(filename),
                        FullSizeImageFile = filename,
                        ContextLink = filename,
                        Tooltip = ImageUtil.GetWidthAndHeight(filename)
                    };
                    //add it
                    ImageSelectionList.Add(fromFileItem);
                }
            }
            catch
            {
                // not an image or image is corrupted
                ImagesLabWindow.ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorImageCorrupted);
            }
        }

        public void AddImageSelectionListItem(string downloadLink)
        {
            if (StringUtil.IsEmpty(downloadLink) || !UrlUtil.IsUrlValid(downloadLink)) // Case 1: If url not valid
            {
                ImagesLabWindow.ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorUrlLinkIncorrect);
                return;
            }
            var item = new ImageItem
            {
                ImageFile = StoragePath.LoadingImgPath,
                ContextLink = downloadLink
            };
            UrlUtil.GetMetaInfo(ref downloadLink, item);
            ImageSelectionList.Add(item);
            ImagesLabWindow.ActivateImageDownloadProgressRing();

            var imagePath = StoragePath.GetPath("img-"
                + DateTime.Now.GetHashCode() + "-"
                + Guid.NewGuid().ToString().Substring(0, 7));
            ImageDownloader
                .Get(downloadLink, imagePath)
                .After(() =>
                {
                    try
                    {
                        VerifyIsProperImage(imagePath); // Case 2: not a proper image
                        item.UpdateDownloadedImage(imagePath);
                        ImagesLabWindow.UpdatePreviewImagesForDownloadedImage(item);
                    }
                    catch
                    {
                        ImagesLabWindow.ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorImageDownloadCorrupted);
                        ImageSelectionList.Remove(item);
                    }
                    finally
                    {
                        ImagesLabWindow.DeactivateImageDownloadProgressRing();
                    }
                })
                // Case 3: Possibly network timeout
                .OnError(e =>
                {
                    ImagesLabWindow.DeactivateImageDownloadProgressRing();
                    ImageSelectionList.Remove(item);
                    ImagesLabWindow.ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorFailedToLoad + e.Message);
                })
                .Start();
        }

        private static void VerifyIsProperImage(string filename)
        {
            using (Image.FromFile(filename))
            {
                // so this is a proper image
            }
        }

        public void UpdatePreviewImages(ImageItem source)
        {
            StylesPreviewList.Clear();
            if (PowerPointCurrentPresentationInfo.CurrentSlide == null) return;
            try
            {
                foreach (var stylesPreviewOption in StyleOptionsFactory.GetAllStylesPreviewOptions())
                {
                    var previewInfo = PreviewPresentation.PreviewApplyStyle(source, stylesPreviewOption);
                    StylesPreviewList.Add(new ImageItem
                    {
                        ImageFile = previewInfo.PreviewApplyStyleImagePath,
                        Tooltip = stylesPreviewOption.StyleName
                    });
                }
            }
            catch
            {
                ImagesLabWindow.ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorImageCorrupted);
            }
        }

        public void ApplyStyleInPreviewStage(ImageItem source, string targetStyle)
        {
            try
            {
                var targetDefaultOptions = StyleOptionsFactory.GetStylesPreviewOption(targetStyle);
                PreviewPresentation.ApplyStyle(source, targetDefaultOptions);
                ImagesLabWindow.ShowSuccessfullyAppliedDialog();
            }
            catch (AssumptionFailedException)
            {
                ImagesLabWindow.ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorNoSelectedSlide);
            }
        }

        public void InitStyleVariationCategories(IList<StyleOptions> givenOptions,
            Dictionary<string, List<StyleVariants>> givenVariants, string targetStyle)
        {
            _styleOptions = givenOptions ?? StyleOptionsFactory.GetStylesVariationOptions(targetStyle);
            _styleVariants = givenVariants ?? StyleVariantsFactory.GetVariants(targetStyle);
            _previousVariantsCategory = ImagesLabWindow.InitVariantsComboBox(_styleVariants);

            // default style options (in preview stage)
            var defaultStyleOptions = StyleOptionsFactory.GetStylesPreviewOption(targetStyle);
            var currentVariants = _styleVariants.Values.First();
            var variantIndexWithoutEffect = -1;
            for (var i = 0; i < currentVariants.Count; i++)
            {
                if (currentVariants[i].IsNoEffect(defaultStyleOptions))
                {
                    variantIndexWithoutEffect = i;
                    break;
                }
            }

            // swap the no-effect variant with the current selected style's corresponding variant
            // so that to achieve continuity.
            // in order to swap, style option provided from StyleOptionsFactory should have
            // corresponding values specified in StyleVariantsFactory. e.g., an option generated
            // from factory has overlay transparency of 35, then in order to swap, it should have
            // a variant of overlay transparency of 35. Otherwise it cannot swap, because variants
            // don't match any values in the style options.
            if (variantIndexWithoutEffect != -1 && givenOptions == null)
            {
                // swap style variant
                var tempVariant = currentVariants[variantIndexWithoutEffect];
                currentVariants[variantIndexWithoutEffect] =
                    currentVariants[0];
                currentVariants[0] = tempVariant;
                // swap default style options (in variation stage)
                var tempStyleOpt = _styleOptions[variantIndexWithoutEffect];
                _styleOptions[variantIndexWithoutEffect] =
                    _styleOptions[0];
                _styleOptions[0] = tempStyleOpt;
            }

            for (var i = 0; i < currentVariants.Count && i < _styleOptions.Count; i++)
            {
                currentVariants[i].Apply(_styleOptions[i]);
            }
        }

        public void UpdateStyleVariationImages(ImageItem source)
        {
            try
            {
                var scrollOffset = ImagesLabWindow.GetVariationListBoxScrollOffset();
                var selectedId = ImagesLabWindow.GetVariationListBoxSelectedId();
                ClearStyleVariationList();
                foreach (var styleOption in _styleOptions)
                {
                    var previewInfo = PreviewPresentation.PreviewApplyStyle(source, styleOption);
                    StylesVariationList.Add(new ImageItem
                    {
                        ImageFile = previewInfo.PreviewApplyStyleImagePath,
                        Tooltip = styleOption.OptionName
                    });
                }
                ImagesLabWindow.SetVariationListBoxSelectedId(selectedId);
                ImagesLabWindow.SetVariationListBoxScrollOffset(scrollOffset);
            }
            catch
            {
                ImagesLabWindow.ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorImageCorrupted);
            }
        }

        public void UpdateStyleVariationCategories()
        {
            var targetVariants = _styleVariants[_previousVariantsCategory];
            if (targetVariants.Count == 0) return;

            var targetVariationSelectedIndex = ImagesLabWindow.GetVariationListBoxSelectedId();
            var targetVariant = targetVariants[targetVariationSelectedIndex];
            foreach (var option in _styleOptions)
            {
                targetVariant.Apply(option);
            }

            var currentVariantsCategory = ImagesLabWindow.GetVariantsComboBoxSelectedValue();
            if (currentVariantsCategory != TextCollection.ImagesLabText.VariantCategoryFontColor
                && _previousVariantsCategory != TextCollection.ImagesLabText.VariantCategoryFontColor)
            {
                // apply font color variant,
                // because default styles may contain special font color settings, but not in variants
                var fontColorVariant = new StyleVariants(new Dictionary<string, object>
                {
                    {"FontColor", _styleOptions[targetVariationSelectedIndex].FontColor}
                });
                foreach (var option in _styleOptions)
                {
                    fontColorVariant.Apply(option);
                }
            }

            var nextCategoryVariants = _styleVariants[currentVariantsCategory];
            var variantIndexWithoutEffect = -1;
            for (var i = 0; i < nextCategoryVariants.Count; i++)
            {
                if (nextCategoryVariants[i].IsNoEffect(_styleOptions[targetVariationSelectedIndex]))
                {
                    variantIndexWithoutEffect = i;
                    break;
                }
            }
            // swap the no-effect variant with the current selected style's corresponding variant
            // so that to achieve an effect: jumpt between different category wont change the
            // selected style
            if (variantIndexWithoutEffect != -1)
            {
                var temp = nextCategoryVariants[variantIndexWithoutEffect];
                nextCategoryVariants[variantIndexWithoutEffect] =
                    nextCategoryVariants[targetVariationSelectedIndex];
                nextCategoryVariants[targetVariationSelectedIndex] = temp;
            }

            for (var i = 0; i < nextCategoryVariants.Count && i < _styleOptions.Count; i++)
            {
                nextCategoryVariants[i].Apply(_styleOptions[i]);
            }

            _previousVariantsCategory = currentVariantsCategory;
            ImagesLabWindow.UpdateStyleVariationsImages();
        }

        public StyleOptions GetStyleVariationStyleOptions(int index)
        {
            return _styleOptions[index];
        }

        public StyleVariants GetStyleVariationStyleVariants(string category, int index)
        {
            return _styleVariants[category][index];
        }

        public void ApplyStyleInVariationStage(ImageItem source)
        {
            try
            {
                PreviewPresentation.ApplyStyle(source);
                ImagesLabWindow.ShowSuccessfullyAppliedDialog();
            }
            catch (AssumptionFailedException)
            {
                ImagesLabWindow.ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorNoSelectedSlide);
            }
        }

        public void SetStyleDesignerOptions(StyleOptions option)
        {
            PreviewPresentation.SetStyleOptions(option);
        }
    }
}
