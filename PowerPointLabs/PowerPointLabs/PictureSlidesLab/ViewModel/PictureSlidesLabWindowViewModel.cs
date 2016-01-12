using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.AutoUpdate.Interface;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory;
using PowerPointLabs.PictureSlidesLab.Service;
using PowerPointLabs.PictureSlidesLab.Service.Interface;
using PowerPointLabs.PictureSlidesLab.Thread;
using PowerPointLabs.PictureSlidesLab.Util;
using PowerPointLabs.PictureSlidesLab.View.Interface;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Exceptions;
using PowerPointLabs.WPF.Observable;

namespace PowerPointLabs.PictureSlidesLab.ViewModel
{
    partial class PictureSlidesLabWindowViewModel
    {
        #region UI Models
        // UI model - for image selection stage
        public ObservableCollection<ImageItem> ImageSelectionList { get; set; }

        public ObservableInt ImageSelectionListSelectedId { get; set; }

        public ObservableImageItem ImageSelectionListSelectedItem { get; set; }

        public ObservableBoolean IsActiveDownloadProgressRing { get; set; }

        // UI model - for preview stage
        public ObservableCollection<ImageItem> StylesPreviewList { get; set; }

        public ObservableInt StylesPreviewListSelectedId { get; set; }

        public ObservableImageItem StylesPreviewListSelectedItem { get; set; }

        // UI model - for variation stage
        public ObservableCollection<ImageItem> StylesVariationList { get; set; }

        public ObservableInt StylesVariationListSelectedId { get; set; }

        public ObservableImageItem StylesVariationListSelectedItem { get; set; }

        public ObservableString CurrentVariantCategory { get; set; }

        public ObservableInt CurrentVariantCategoryId { get; set; }

        public ObservableCollection<string> VariantsCategory { get; set; }

        public ObservableInt SelectedFontId { get; set; }

        public ObservableFont SelectedFontFamily { get; set; }

        #endregion

        #region Dependency

        // UI controller
        public IPictureSlidesLabWindowView View { private get; set; }

        // Downloader
        public IDownloader ImageDownloader { private get; set; }

        // Background presentation that will do the styles processing
        public IStylesDesigner Designer { private get; set; }

        #endregion

        #region States for variation stage
        private string _previousVariantsCategory;
        private List<StyleOptions> _styleOptions;
        // key - variant category, value - variants
        private Dictionary<string, List<StyleVariants>> _styleVariants;
        #endregion

        #region Lifecycle
        public PictureSlidesLabWindowViewModel(IPictureSlidesLabWindowView view)
        {
            View = view;
            ImageDownloader = new ContextDownloader(View.GetThreadContext());
            InitUiModels();
            InitStorage();
            Designer = new StylesDesigner();
        }

        private void InitUiModels()
        {
            StylesVariationList = new ObservableCollection<ImageItem>();
            StylesVariationListSelectedId = new ObservableInt {Number = -1};
            StylesVariationListSelectedItem = new ObservableImageItem();
            CurrentVariantCategory = new ObservableString();
            CurrentVariantCategoryId = new ObservableInt {Number = -1};
            VariantsCategory = new ObservableCollection<string>();
            SelectedFontId = new ObservableInt();
            SelectedFontFamily = new ObservableFont();

            StylesPreviewList = new ObservableCollection<ImageItem>();
            StylesPreviewListSelectedId = new ObservableInt {Number = -1};
            StylesPreviewListSelectedItem = new ObservableImageItem();

            ImageSelectionList = StoragePath.Load();
            ImageSelectionListSelectedId = new ObservableInt {Number = -1};
            ImageSelectionListSelectedItem = new ObservableImageItem();
            IsActiveDownloadProgressRing = new ObservableBoolean {Flag = false};
        }

        private void InitStorage()
        {
            var imageFilesInUse = new HashSet<string>();
            foreach (var imageItem in ImageSelectionList)
            {
                imageFilesInUse.Add(imageItem.ImageFile);
                imageFilesInUse.Add(imageItem.FullSizeImageFile);
                if (imageItem.CroppedImageFile != null)
                {
                    imageFilesInUse.Add(imageItem.CroppedImageFile);
                    imageFilesInUse.Add(imageItem.CroppedThumbnailImageFile);
                }
            }

            var isTempPathInit = TempPath.InitTempFolder();
            var isStoragePathInit = StoragePath.InitPersistentFolder(imageFilesInUse);
            if (!isTempPathInit || !isStoragePathInit)
            {
                View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorFailToInitTempFolder);
            }
        }

        public void CleanUp()
        {
            if (Designer != null)
            {
                Designer.CleanUp();
            }
            StoragePath.Save(ImageSelectionList);
        }
        #endregion

        #region Stage - Image Selection (Add Image)
        /// <summary>
        /// Add images from local files
        /// </summary>
        /// <param name="filenames"></param>
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
                View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorImageCorrupted);
            }
        }

        /// <summary>
        /// Add image by downloading
        /// </summary>
        /// <param name="downloadLink"></param>
        /// <param name="contentSlide"></param>
        /// <param name="slideWidth"></param>
        /// <param name="slideHeight"></param>
        public void AddImageSelectionListItem(string downloadLink, 
            Slide contentSlide, float slideWidth, float slideHeight)
        {
            if (StringUtil.IsEmpty(downloadLink) || !UrlUtil.IsUrlValid(downloadLink)) // Case 1: If url not valid
            {
                View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorUrlLinkIncorrect);
                return;
            }
            var item = new ImageItem
            {
                ImageFile = StoragePath.LoadingImgPath,
                ContextLink = downloadLink
            };
            UrlUtil.GetMetaInfo(ref downloadLink, item);
            ImageSelectionList.Add(item);
            IsActiveDownloadProgressRing.Flag = true;

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
                        if (ImageSelectionListSelectedItem.ImageItem != null 
                            && imagePath == ImageSelectionListSelectedItem.ImageItem.ImageFile)
                        {
                            UpdatePreviewImages(contentSlide, slideWidth, slideHeight);
                        }
                    }
                    catch
                    {
                        View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorImageDownloadCorrupted);
                        ImageSelectionList.Remove(item);
                    }
                    finally
                    {
                        IsActiveDownloadProgressRing.Flag = false;
                    }
                })
                // Case 3: Possibly network timeout
                .OnError(e =>
                {
                    IsActiveDownloadProgressRing.Flag = false;
                    ImageSelectionList.Remove(item);
                    View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorFailedToLoad + e.Message);
                })
                .Start();
        }
        #endregion

        #region Stage - Styles Previewing
        public void UpdatePreviewImages(Slide contentSlide, float slideWidth, float slideHeight)
        {
            if (View.IsVariationsFlyoutOpen)
            {
                UpdateStylesVariationImages(contentSlide, slideWidth, slideHeight);
            }
            else
            {
                UpdateStylesPreviewImages(contentSlide, slideWidth, slideHeight);
            }
        }

        public void ApplyStyleInPreviewStage(Slide contentSlide, float slideWidth, float slideHeight)
        {
            try
            {
                var targetDefaultOptions = StyleOptionsFactory
                    .GetStylesPreviewOption(StylesPreviewListSelectedItem.ImageItem.Tooltip);
                Designer.ApplyStyle(ImageSelectionListSelectedItem.ImageItem, contentSlide,
                    slideWidth, slideHeight, targetDefaultOptions);
                View.ShowSuccessfullyAppliedDialog();
            }
            catch (AssumptionFailedException)
            {
                View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorNoSelectedSlide);
            }
        }
        #endregion

        #region Stage - Styles Variation
        /// <summary>
        /// When stylesVariationListSelectedItem is changed,
        /// this method will be called to update the corresponding style options of designer
        /// </summary>
        public void UpdateStyleVariationStyleOptionsWhenSelectedItemChange()
        {
            Designer.SetStyleOptions(_styleOptions[StylesVariationListSelectedId.Number]);
        }

        /// <summary>
        /// Update styles variation iamges before its flyout is open
        /// </summary>
        /// <param name="contentSlide"></param>
        /// <param name="slideWidth"></param>
        /// <param name="slideHeight"></param>
        /// <param name="givenOptions"></param>
        /// <param name="givenVariants"></param>
        public void UpdateStyleVariationImagesWhenOpenFlyout(Slide contentSlide, float slideWidth, float slideHeight,
            List<StyleOptions> givenOptions = null, Dictionary<string, List<StyleVariants>> givenVariants = null)
        {
            var targetStyleItem = StylesPreviewListSelectedItem.ImageItem;
            var source = ImageSelectionListSelectedItem.ImageItem;
            StylesVariationList.Clear();

            if (!IsAbleToUpdateStylesVariationImages(source, targetStyleItem, contentSlide))
                return;

            InitStylesVariationCategories(givenOptions, givenVariants, targetStyleItem.Tooltip);
            UpdateStylesVariationImages(source, contentSlide, slideWidth, slideHeight);

            StylesVariationListSelectedId.Number = 0;
            View.SetVariationListBoxScrollOffset(0d);
        }

        /// <summary>
        /// Update styles variation images after its flyout been open
        /// </summary>
        public void UpdateStylesVariationImages(Slide contentSlide, float slideWidth, float slideHeight)
        {
            var selectedId = StylesVariationListSelectedId.Number;
            var scrollOffset = View.GetVariationListBoxScrollOffset();
            var targetStyleItem = StylesPreviewListSelectedItem.ImageItem;
            var source = ImageSelectionListSelectedItem.ImageItem;
            StylesVariationList.Clear();

            if (!IsAbleToUpdateStylesVariationImages(source, targetStyleItem, contentSlide))
                return;

            UpdateStylesVariationImages(source, contentSlide, slideWidth, slideHeight);

            StylesVariationListSelectedId.Number = selectedId;
            View.SetVariationListBoxScrollOffset(scrollOffset);
        }

        /// <summary>
        /// This method implements the way to guide the user step by step to customize
        /// style
        /// </summary>
        public void UpdateStepByStepStylesVariationImages(Slide contentSlide, float slideWidth, float slideHeight)
        {
            if (StylesVariationListSelectedId.Number < 0
                || VariantsCategory.Count == 0) return;

            var targetVariationSelectedIndex = StylesVariationListSelectedId.Number;
            var targetVariant = _styleVariants[_previousVariantsCategory][targetVariationSelectedIndex];
            foreach (var option in _styleOptions)
            {
                targetVariant.Apply(option);
            }
            
            var currentVariantsCategory = CurrentVariantCategory.Text;
            if (currentVariantsCategory != TextCollection.PictureSlidesLabText.VariantCategoryFontColor
                && _previousVariantsCategory != TextCollection.PictureSlidesLabText.VariantCategoryFontColor)
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
            UpdateStylesVariationImages(contentSlide, slideWidth, slideHeight);
        }

        public void ApplyStyleInVariationStage(Slide contentSlide, float slideWidth, float slideHeight)
        {
            try
            {
                Designer.ApplyStyle(ImageSelectionListSelectedItem.ImageItem, contentSlide,
                    slideWidth, slideHeight);
                View.ShowSuccessfullyAppliedDialog();
            }
            catch (AssumptionFailedException)
            {
                View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorNoSelectedSlide);
            }
        }
        #endregion

        #region Helper funcs
        private static void VerifyIsProperImage(string filename)
        {
            using (Image.FromFile(filename))
            {
                // so this is a proper image
            }
        }

        private void UpdateStylesPreviewImages(Slide contentSlide, float slideWidth, float slideHeight)
        {
            var selectedId = StylesPreviewListSelectedId.Number;
            var source = ImageSelectionListSelectedItem.ImageItem;
            StylesPreviewList.Clear();

            if (!IsAbleToUpdateStylesPreviewImages(source, contentSlide))
                return;

            try
            {
                foreach (var stylesPreviewOption in StyleOptionsFactory.GetAllStylesPreviewOptions())
                {
                    var previewInfo = Designer.PreviewApplyStyle(source, 
                        contentSlide, slideWidth, slideHeight, stylesPreviewOption);
                    StylesPreviewList.Add(new ImageItem
                    {
                        ImageFile = previewInfo.PreviewApplyStyleImagePath,
                        Tooltip = stylesPreviewOption.StyleName
                    });
                }
            }
            catch
            {
                View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorImageCorrupted);
            }

            StylesPreviewListSelectedId.Number = selectedId;
        }

        private static bool IsAbleToUpdateStylesPreviewImages(ImageItem source, Slide contentSlide)
        {
            return !(source == null
                    || source.ImageFile == StoragePath.LoadingImgPath
                    || contentSlide == null);
        }

        private void InitStylesVariationCategories(List<StyleOptions> givenOptions,
            Dictionary<string, List<StyleVariants>> givenVariants, string targetStyle)
        {
            _styleOptions = givenOptions ?? StyleOptionsFactory.GetStylesVariationOptions(targetStyle);
            _styleVariants = givenVariants ?? StyleVariantsFactory.GetVariants(targetStyle);

            VariantsCategory.Clear();
            foreach (var styleVariant in _styleVariants.Keys)
            {
                VariantsCategory.Add(styleVariant);
            }
            CurrentVariantCategoryId.Number = 0;
            _previousVariantsCategory = VariantsCategory[0];

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

        private static bool IsAbleToUpdateStylesVariationImages(ImageItem source, ImageItem targetStyleItem, 
            Slide contentSlide)
        {
            return !(source == null
                    || source.ImageFile == StoragePath.LoadingImgPath
                    || targetStyleItem == null
                    || targetStyleItem.Tooltip == null
                    || contentSlide == null);
        }

        private void UpdateStylesVariationImages(ImageItem source, Slide contentSlide, 
            float slideWidth, float slideHeight)
        {
            try
            {
                foreach (var styleOption in _styleOptions)
                {
                    var previewInfo = Designer.PreviewApplyStyle(source, contentSlide, 
                        slideWidth, slideHeight, styleOption);
                    StylesVariationList.Add(new ImageItem
                    {
                        ImageFile = previewInfo.PreviewApplyStyleImagePath,
                        Tooltip = styleOption.OptionName
                    });
                }
            }
            catch
            {
                View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorImageCorrupted);
            }
        }
        #endregion
    }
}
