using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.AutoUpdate.Interface;
using PowerPointLabs.Models;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory;
using PowerPointLabs.PictureSlidesLab.Service;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.PictureSlidesLab.Service.Interface;
using PowerPointLabs.PictureSlidesLab.Service.Preview;
using PowerPointLabs.PictureSlidesLab.Thread;
using PowerPointLabs.PictureSlidesLab.Util;
using PowerPointLabs.PictureSlidesLab.View.Interface;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Exceptions;
using PowerPointLabs.WPF.Observable;
using Fonts = System.Windows.Media.Fonts;

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

        public ObservableCollection<string> FontFamilies { get; set; } 

        public ObservableInt SelectedFontId { get; set; }

        public ObservableFont SelectedFontFamily { get; set; }

        public Settings Settings { get; set; }

        #endregion

        #region Dependency

        // UI controller
        public IPictureSlidesLabWindowView View { private get; set; }

        // Downloader
        public IDownloader ImageDownloader { private get; set; }

        // Background presentation that will do the styles processing
        public IStylesDesigner Designer { private get; set; }

        // Style Options Factory
        public StyleOptionsFactory OptionsFactory { get; set; }

        // Style Variants Factory
        public StyleVariantsFactory VariantsFactory { get; set; }

        #endregion

        #region States for variation stage
        private string _previousVariantsCategory;
        private List<StyleOption> _styleOptions;
        // key - variant category, value - variants
        private Dictionary<string, List<StyleVariant>> _styleVariants;
        // for picture variation
        private List<ImageItem> _8PicturesInPictureVariation;
        private bool _isPictureVariationInit;
        #endregion

        #region Lifecycle
        public PictureSlidesLabWindowViewModel(IPictureSlidesLabWindowView view, 
            IStylesDesigner stylesDesigner = null)
        {
            View = view;
            ImageDownloader = new ContextDownloader(View.GetThreadContext());
            InitStorage();
            InitUiModels();
            InitFontFamilies();
            CleanUnusedPersistentData();
            Designer = stylesDesigner ?? new StylesDesigner();
            Designer.SetSettings(Settings);
            OptionsFactory = new StyleOptionsFactory();
            VariantsFactory = new StyleVariantsFactory();
        }

        private void InitFontFamilies()
        {
            FontFamilies = new ObservableCollection<string>();
            foreach (var fontFamily in Fonts.SystemFontFamilies)
            {
                FontFamilies.Add(fontFamily.Source);
            }

            // add font family not in Fonts.SystemFontFamilies
            FontFamilies.Add("Segoe UI Light");
            FontFamilies.Add("Calibri Light");
            FontFamilies.Add("Arial Black");
            FontFamilies.Add("Times New Roman Italic");
        }

        private void CleanUnusedPersistentData()
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
            StoragePath.CleanPersistentFolder(imageFilesInUse);
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

            ImageSelectionList = new ObservableCollection<ImageItem>();
            ImageSelectionList.Add(CreateChoosePicturesItem());

            Settings = StoragePath.LoadSettings();

            if (StoragePath.IsFirstTimeUsage())
            {
                ImageSelectionList.Add(CreateSamplePic1Item());
                ImageSelectionList.Add(CreateSamplePic2Item());
            }
            else
            {
                var loadedImageSelectionList = StoragePath.LoadPictures();
                foreach (var item in loadedImageSelectionList)
                {
                    ImageSelectionList.Add(item);
                }
            }

            ImageSelectionListSelectedId = new ObservableInt {Number = -1};
            ImageSelectionListSelectedItem = new ObservableImageItem();
            IsActiveDownloadProgressRing = new ObservableBoolean {Flag = false};
        }

        private void InitStorage()
        {
            var isTempPathInit = Util.TempPath.InitTempFolder();
            var isStoragePathInit = StoragePath.InitPersistentFolder();
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
            ImageSelectionList.RemoveAt(0);
            StoragePath.Save(ImageSelectionList);
            StoragePath.Save(Settings);
        }
        #endregion

        #region Stage - Image Selection (Add Image)

        public void RemoveAllImageSelectionListItems()
        {
            ImageSelectionList.Clear();
            ImageSelectionList.Add(CreateChoosePicturesItem());
        }

        /// <summary>
        /// Add image from local files
        /// </summary>
        /// <param name="filenames"></param>
        /// <param name="contentSlide"></param>
        /// <param name="slideWidth"></param>
        /// <param name="slideHeight"></param>
        public void AddImageSelectionListItem(string[] filenames, Slide contentSlide, 
            float slideWidth, float slideHeight)
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
                        Source = "local drive",
                        Tooltip = ImageUtil.GetWidthAndHeight(filename)
                    };
                    //add it
                    ImageSelectionList.Add(fromFileItem);
                    UpdatePictureInPictureVariationWhenAddedNewOne(fromFileItem);
                }
                if (IsInPictureVariation() && filenames.Length > 0)
                {
                    UpdatePreviewImages(
                        ImageSelectionListSelectedItem.ImageItem ?? View.CreateDefaultPictureItem(),
                        contentSlide, slideWidth, slideHeight);
                }
            }
            catch (Exception e)
            {
                // not an image or image is corrupted
                View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorImageCorrupted);
                Logger.LogException(e, "AddImageSelectionListItem");
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
                ContextLink = downloadLink,
                Source = downloadLink
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
                        UpdatePictureInPictureVariationWhenAddedNewOne(item);
                        if (ImageSelectionListSelectedItem.ImageItem != null
                            && imagePath == ImageSelectionListSelectedItem.ImageItem.FullSizeImageFile)
                        {
                            UpdatePreviewImages(ImageSelectionListSelectedItem.ImageItem,
                                contentSlide, slideWidth, slideHeight);
                        }
                        else if (IsInPictureVariation())
                        {
                            UpdatePreviewImages(
                                ImageSelectionListSelectedItem.ImageItem ?? View.CreateDefaultPictureItem(),
                                contentSlide, slideWidth, slideHeight);
                        }
                    }
                    catch (Exception e)
                    {
                        View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorImageDownloadCorrupted);
                        ImageSelectionList.Remove(item);
                        Logger.LogException(e, "AddImageSelectionListItem (download)");
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

        /// <summary>
        /// General update preview images,
        /// can be used in most use cases, such as reload preview images
        /// after re-activate PSL main window.
        /// </summary>
        /// <param name="source"></param>
        /// <param name="contentSlide"></param>
        /// <param name="slideWidth"></param>
        /// <param name="slideHeight"></param>
        public void UpdatePreviewImages(ImageItem source, Slide contentSlide, float slideWidth, float slideHeight)
        {
            if (View.IsVariationsFlyoutOpen)
            {
                UpdateStylesVariationImagesAfterOpenFlyout(source, contentSlide, slideWidth, slideHeight);
            }
            else
            {
                UpdateStylesPreviewImages(source, contentSlide, slideWidth, slideHeight);
            }
        }

        public void ApplyStyleInPreviewStage(Slide contentSlide, float slideWidth, float slideHeight)
        {
            var copiedPicture = LoadClipboardPicture();
            try
            {
                var targetDefaultOptions = OptionsFactory
                    .GetStylesPreviewOption(StylesPreviewListSelectedItem.ImageItem.Tooltip);
                Designer.ApplyStyle(ImageSelectionListSelectedItem.ImageItem, contentSlide,
                    slideWidth, slideHeight, targetDefaultOptions);
                View.ShowSuccessfullyAppliedDialog();
            }
            catch (Exception e)
            {
                View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorNoSelectedSlide);
                Logger.LogException(e, "ApplyStyleInPreviewStage");
            }
            SaveClipboardPicture(copiedPicture);
        }
        #endregion

        #region Stage - Styles Variation
        /// <summary>
        /// When stylesVariationListSelectedItem is changed,
        /// this method will be called to update the corresponding style options of designer
        /// </summary>
        public void UpdateStyleVariationStyleOptionsWhenSelectedItemChange()
        {
            try
            {
                Designer.SetStyleOptions(_styleOptions[StylesVariationListSelectedId.Number]);
            }
            catch (Exception e)
            {
                View.ShowErrorMessageBox("Failed when retrieving information from the selected preview image.", e);
                Logger.LogException(e, "UpdateStyleVariationStyleOptionsWhenSelectedItemChange");
            }
        }

        /// <summary>
        /// Update styles variation iamges before its flyout is open
        /// </summary>
        /// <param name="source"></param>
        /// <param name="contentSlide"></param>
        /// <param name="slideWidth"></param>
        /// <param name="slideHeight"></param>
        /// <param name="givenOptions"></param>
        /// <param name="givenVariants"></param>
        public void UpdateStyleVariationImagesWhenOpenFlyout(ImageItem source, Slide contentSlide, float slideWidth, float slideHeight,
            List<StyleOption> givenOptions = null, Dictionary<string, List<StyleVariant>> givenVariants = null)
        {
            Logger.Log("Variation stage is open");
            var targetStyleItem = StylesPreviewListSelectedItem.ImageItem;
            StylesVariationList.Clear();

            if (!IsAbleToUpdateStylesVariationImages(source, targetStyleItem, contentSlide))
                return;

            InitStylesVariationCategories(givenOptions, givenVariants, targetStyleItem.Tooltip);
            if (Settings.GetDefaultAspectWhenCustomize() == Aspect.PictureAspect)
            {
                UpdateStylesVariationImages(source, contentSlide, slideWidth, slideHeight, isMockPreviewImages: true);
            }
            else
            {
                UpdateStylesVariationImages(source, contentSlide, slideWidth, slideHeight);
            }

            StylesVariationListSelectedId.Number = 0;
            View.SetVariationListBoxScrollOffset(0d);
            _isPictureVariationInit = false;

            if (Settings.GetDefaultAspectWhenCustomize() == Aspect.PictureAspect)
            {
                CurrentVariantCategoryId.Number =
                    VariantsCategory.IndexOf(TextCollection.PictureSlidesLabText.VariantCategoryPicture);
            }
        }

        /// <summary>
        /// Update styles variation images after its flyout been open
        /// </summary>
        public void UpdateStylesVariationImagesAfterOpenFlyout(ImageItem source, Slide contentSlide, float slideWidth, float slideHeight)
        {
            var selectedId = StylesVariationListSelectedId.Number;
            var scrollOffset = View.GetVariationListBoxScrollOffset();
            var targetStyleItem = StylesPreviewListSelectedItem.ImageItem;
            StylesVariationList.Clear();

            if (!IsAbleToUpdateStylesVariationImages(source, targetStyleItem, contentSlide))
                return;

            UpdateStylesVariationImages(source, contentSlide, slideWidth, slideHeight);

            StylesVariationListSelectedId.Number = selectedId < 0 ? 0 : selectedId;
            View.SetVariationListBoxScrollOffset(scrollOffset);
        }

        /// <summary>
        /// This method implements the way to guide the user step by step to customize
        /// style
        /// </summary>
        public void UpdateStepByStepStylesVariationImages(ImageItem source, Slide contentSlide, 
            float slideWidth, float slideHeight)
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
                var fontColorVariant = new StyleVariant(new Dictionary<string, object>
                {
                    {"FontColor", _styleOptions[targetVariationSelectedIndex].FontColor}
                });
                foreach (var option in _styleOptions)
                {
                    fontColorVariant.Apply(option);
                }
            }

            var nextCategoryVariants = _styleVariants[currentVariantsCategory];
            if (currentVariantsCategory == TextCollection.PictureSlidesLabText.VariantCategoryFontFamily)
            {
                var isFontInVariation = false;
                var currentFontFamily = _styleOptions[targetVariationSelectedIndex].FontFamily;
                foreach (var variant in nextCategoryVariants)
                {
                    if (currentFontFamily == (string) variant.Get("FontFamily"))
                    {
                        isFontInVariation = true;
                    }
                }
                if (!isFontInVariation 
                    && targetVariationSelectedIndex >= 0 
                    && targetVariationSelectedIndex < nextCategoryVariants.Count)
                {
                    nextCategoryVariants[targetVariationSelectedIndex]
                        .Set("FontFamily", currentFontFamily);
                    nextCategoryVariants[targetVariationSelectedIndex]
                        .Set("OptionName", currentFontFamily);
                }
            }
            
            int pictureIndexToSelect = -1;
            // Enter picture variation for the first time
            if (CurrentVariantCategory.Text == TextCollection.PictureSlidesLabText.VariantCategoryPicture
                && !_isPictureVariationInit)
            {
                _8PicturesInPictureVariation = GetLast8Pictures(targetVariationSelectedIndex);
                _isPictureVariationInit = true;
            }
            // Enter picture variation again
            else if (CurrentVariantCategory.Text == TextCollection.PictureSlidesLabText.VariantCategoryPicture
                     && _isPictureVariationInit)
            {
                var isPictureSwapped = false;
                for (var i = 0; i < _8PicturesInPictureVariation.Count; i++)
                {
                    // swap the picture to the current selected id in
                    // variation list
                    var picture = _8PicturesInPictureVariation[i];
                    if ((ImageSelectionListSelectedItem.ImageItem == null 
                        && picture.ImageFile == StoragePath.NoPicturePlaceholderImgPath) || 
                            (ImageSelectionListSelectedItem.ImageItem != null
                            && picture.ImageFile == ImageSelectionListSelectedItem.ImageItem.ImageFile))
                    {
                        var tempPic = _8PicturesInPictureVariation[targetVariationSelectedIndex];
                        _8PicturesInPictureVariation[targetVariationSelectedIndex]
                            = picture;
                        _8PicturesInPictureVariation[i] = tempPic;
                        isPictureSwapped = true;
                        break;
                    }
                }
                if (!isPictureSwapped)
                {
                    // if the current picture doesn't exist in the _8PicturesInPictureVariation
                    // directly overwrite the existing one at the selected id
                    UpdateSelectedPictureInPictureVariation();
                }
            }
            // Exit picture variation
            else if (_previousVariantsCategory == TextCollection.PictureSlidesLabText.VariantCategoryPicture)
            {
                // use the selected picture in the picture variation to preview
                var targetPicture = _8PicturesInPictureVariation[targetVariationSelectedIndex];
                if (targetPicture.ImageFile != StoragePath.NoPicturePlaceholderImgPath)
                {
                    var indexForTargetPicture = ImageSelectionList.IndexOf(targetPicture);
                    if (indexForTargetPicture == -1)
                    {
                        ImageSelectionList.Add(targetPicture);
                        pictureIndexToSelect = ImageSelectionList.Count - 1;
                    }
                    else
                    {
                        pictureIndexToSelect = indexForTargetPicture;
                    }
                }
                else // target picture is the default picture
                {
                    // enter default picture mode
                    View.DisableUpdatingPreviewImages();
                    ImageSelectionListSelectedId.Number = -1;
                    source = View.CreateDefaultPictureItem();
                }
            }

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
            if (pictureIndexToSelect == -1
                || pictureIndexToSelect == ImageSelectionListSelectedId.Number)
            {
                UpdateStylesVariationImagesAfterOpenFlyout(source, contentSlide,
                    slideWidth, slideHeight);
            }
            else
            {
                ImageSelectionListSelectedId.Number = pictureIndexToSelect;
            }
        }

        public void ApplyStyleInVariationStage(Slide contentSlide, float slideWidth, float slideHeight)
        {
            var copiedPicture = LoadClipboardPicture();
            try
            {
                Designer.ApplyStyle(
                    IsInPictureVariation()
                    ? GetSelectedPictureInPictureVariation(
                        StylesVariationListSelectedId.Number)
                    : ImageSelectionListSelectedItem.ImageItem, contentSlide,
                    slideWidth, slideHeight);
 
                if (IsInPictureVariation())
                {
                    // select the picture if possible
                    var targetPicture = GetSelectedPictureInPictureVariation(
                        StylesVariationListSelectedId.Number);
                    if (targetPicture.ImageFile != StoragePath.NoPicturePlaceholderImgPath)
                    {
                        var indexForTargetPicture = ImageSelectionList.IndexOf(targetPicture);
                        if (indexForTargetPicture == -1)
                        {
                            ImageSelectionList.Add(targetPicture);
                            ImageSelectionListSelectedId.Number = ImageSelectionList.Count - 1;
                        }
                        else
                        {
                            ImageSelectionListSelectedId.Number = indexForTargetPicture;
                        }
                    }
                }
                View.ShowSuccessfullyAppliedDialog();
            }
            catch (Exception e)
            {
                View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorNoSelectedSlide);
                Logger.LogException(e, "ApplyStyleInVariationStage");
            }
            SaveClipboardPicture(copiedPicture);
        }

        #region Picture Variation

        public bool IsInPictureVariation()
        {
            return CurrentVariantCategory != null && CurrentVariantCategory.Text != null
                   && CurrentVariantCategory.Text == TextCollection.PictureSlidesLabText.VariantCategoryPicture;
        }

        public ImageItem GetSelectedPictureInPictureVariation(int pictureIndex)
        {
            try
            {
                return _8PicturesInPictureVariation[pictureIndex];
            }
            catch (Exception e)
            {
                View.ShowErrorMessageBox("Failed when fetching picture aspect.", e);
                Logger.LogException(e, "GetSelectedPictureInPictureVariation");
                return View.CreateDefaultPictureItem();
            }
        }

        public void UpdateSelectedPictureInPictureVariation()
        {
            if (!IsInPictureVariation() 
                || StylesVariationListSelectedId.Number == -1)
                return;

            try
            {
                _8PicturesInPictureVariation[StylesVariationListSelectedId.Number]
                    = ImageSelectionListSelectedItem.ImageItem ?? View.CreateDefaultPictureItem();
            }
            catch (Exception e)
            {
                View.ShowErrorMessageBox("Failed when processing picture aspect.", e);
                Logger.LogException(e, "UpdateSelectedPictureInPictureVariation");
            }
        }

        public void UpdatePictureInPictureVariationWhenAddedNewOne(ImageItem newPicture)
        {
            if (!IsInPictureVariation() || newPicture == null)
                return;

            for (var i = 0; i < _8PicturesInPictureVariation.Count; i++)
            {
                var imageItem = _8PicturesInPictureVariation[i];
                if (imageItem.ImageFile == StoragePath.NoPicturePlaceholderImgPath)
                {
                    _8PicturesInPictureVariation[i] = newPicture;
                    break;
                }
            }
        }

        public void UpdatePictureInPictureVariationWhenDeleteSome()
        {
            if (!IsInPictureVariation())
                return;

            for (var i = 0; i < _8PicturesInPictureVariation.Count; i++)
            {
                var imageItem = _8PicturesInPictureVariation[i];
                if (ImageSelectionList.IndexOf(imageItem) == -1)
                {
                    _8PicturesInPictureVariation[i] = View.CreateDefaultPictureItem();
                }
            }
        }

        private List<ImageItem> GetLast8Pictures(int selectedIdOfVariationList)
        {
            if (!IsInPictureVariation()) return new List<ImageItem>();

            try
            {
                var subPictureList = ImageSelectionList.Skip(Math.Max(1, ImageSelectionList.Count - 8));
                var result = new List<ImageItem>(subPictureList);
                while (result.Count < 8)
                {
                    result.Add(View.CreateDefaultPictureItem());
                }
                if (ImageSelectionListSelectedItem.ImageItem != null
                    && !result.Contains(ImageSelectionListSelectedItem.ImageItem))
                {
                    result[selectedIdOfVariationList] = ImageSelectionListSelectedItem.ImageItem;
                }
                else if (ImageSelectionListSelectedItem.ImageItem == null)
                {
                    for (var i = 0; i < result.Count; i++)
                    {
                        if (result[i].ImageFile == StoragePath.NoPicturePlaceholderImgPath)
                        {
                            result[i] = result[selectedIdOfVariationList];
                            break;
                        }
                    }
                    result[selectedIdOfVariationList] = View.CreateDefaultPictureItem();
                }
                else if (selectedIdOfVariationList >= 0)
                    // contains selected item, need swap to selected index
                {
                    var indexToSwap = result.IndexOf(ImageSelectionListSelectedItem.ImageItem);
                    var tempItem = result[selectedIdOfVariationList];
                    result[selectedIdOfVariationList] = ImageSelectionListSelectedItem.ImageItem;
                    result[indexToSwap] = tempItem;
                }
                return result;
            }
            catch (Exception e)
            {
                View.ShowErrorMessageBox("Failed when generating picture aspect.", e);
                Logger.LogException(e, "GetLast8Pictures");
                return new List<ImageItem>();
            }
        }

        #endregion

        #endregion

        #region Add Picture Citation Slide

        public void AddPictureCitationSlide(Slide slide, List<PowerPointSlide> allSlides)
        {
            new PictureCitationSlide(slide, allSlides).CreatePictureCitations();
        }

        #endregion

        #region Helper funcs

        private static object LoadClipboardPicture()
        {
            try
            {
                var pic = Clipboard.GetImage();
                var text = Clipboard.GetText();
                var files = Clipboard.GetFileDropList();

                if (pic != null)
                {
                    return pic;
                }
                else if (files != null && files.Count > 0)
                {
                    return files;
                }
                else
                {
                    return text;
                }
            }
            catch (Exception e)
            {
                // sometimes Clipboard may fail
                Logger.LogException(e, "LoadClipboardPicture");
                return "";
            }
        }

        private static void SaveClipboardPicture(object copiedObj)
        {
            try
            {
                if (copiedObj is Image)
                {
                    Clipboard.SetImage((Image) copiedObj);
                }
                else if (copiedObj is StringCollection)
                {
                    Clipboard.SetFileDropList((StringCollection) copiedObj);
                }
                else if (!string.IsNullOrEmpty(copiedObj as string))
                {
                    Clipboard.SetText((string) copiedObj);
                }
            }
            catch (Exception e)
            {
                // sometimes Clipboard may fail
                Logger.LogException(e, "SaveClipboardPicture");
            }
        }

        private static void VerifyIsProperImage(string filename)
        {
            using (Image.FromFile(filename))
            {
                // so this is a proper image
            }
        }

        private void UpdateStylesPreviewImages(ImageItem source, Slide contentSlide, float slideWidth, float slideHeight)
        {
            var selectedId = StylesPreviewListSelectedId.Number;
            StylesPreviewList.Clear();

            if (!IsAbleToUpdateStylesPreviewImages(source, contentSlide))
                return;

            var copiedPicture = LoadClipboardPicture();
            try
            {
                foreach (var stylesPreviewOption in OptionsFactory.GetAllStylesPreviewOptions())
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
            catch (Exception e)
            {
                View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorImageCorrupted, e);
                Logger.LogException(e, "UpdateStylesPreviewImages");
            }
            SaveClipboardPicture(copiedPicture);

            StylesPreviewListSelectedId.Number = selectedId < 0 ? 0 : selectedId;
        }

        private static bool IsAbleToUpdateStylesPreviewImages(ImageItem source, Slide contentSlide)
        {
            return !(source == null
                    || source.ImageFile == StoragePath.LoadingImgPath
                    || contentSlide == null);
        }

        private void InitStylesVariationCategories(List<StyleOption> givenOptions,
            Dictionary<string, List<StyleVariant>> givenVariants, string targetStyle)
        {
            _styleOptions = givenOptions ?? OptionsFactory.GetStylesVariationOptions(targetStyle);
            _styleVariants = givenVariants ?? VariantsFactory.GetVariants(targetStyle);

            VariantsCategory.Clear();
            foreach (var styleVariant in _styleVariants.Keys)
            {
                VariantsCategory.Add(styleVariant);
            }
            CurrentVariantCategoryId.Number = 0;
            _previousVariantsCategory = VariantsCategory[0];

            // default style options (in preview stage)
            var defaultStyleOptions = OptionsFactory.GetStylesPreviewOption(targetStyle);
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
            float slideWidth, float slideHeight, bool isMockPreviewImages = false)
        {
            var copiedPicture = LoadClipboardPicture();
            try
            {
                for (var i = 0; i < _styleOptions.Count; i++)
                {
                    var styleOption = _styleOptions[i];
                    PreviewInfo previewInfo;
                    if (isMockPreviewImages)
                    {
                        previewInfo = new PreviewInfo
                        {
                            PreviewApplyStyleImagePath = StoragePath.NoPicturePlaceholderImgPath
                        };
                    }
                    else
                    {
                        previewInfo = Designer.PreviewApplyStyle(
                            IsInPictureVariation()
                                ? _8PicturesInPictureVariation[i]
                                : source,
                            contentSlide, slideWidth, slideHeight, styleOption);
                    }
                    StylesVariationList.Add(new ImageItem
                    {
                        ImageFile = previewInfo.PreviewApplyStyleImagePath,
                        Tooltip = styleOption.OptionName
                    });
                }
            }
            catch (Exception e)
            {
                View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorImageCorrupted, e);
                Logger.LogException(e, "UpdateStylesVariationImages");
            }
            SaveClipboardPicture(copiedPicture);
        }

        private ImageItem CreateChoosePicturesItem()
        {
            return new ImageItem
            {
                ImageFile = StoragePath.ChoosePicturesImgPath,
                Tooltip = "Choose pictures from local storage."
            };
        }

        private static ImageItem CreateSamplePic2Item()
        {
            return new ImageItem
            {
                ImageFile = ImageUtil.GetThumbnailFromFullSizeImg(
                    StoragePath.SampleImg2Path),
                FullSizeImageFile = StoragePath.SampleImg2Path,
                Tooltip = "Picture taken from Gary Elsasser https://flic.kr/p/5s5APp",
                ContextLink = "https://flic.kr/p/5s5APp",
                Source = "https://flic.kr/p/5s5APp"
            };
        }

        private static ImageItem CreateSamplePic1Item()
        {
            return new ImageItem
            {
                ImageFile = ImageUtil.GetThumbnailFromFullSizeImg(
                    StoragePath.SampleImg1Path),
                FullSizeImageFile = StoragePath.SampleImg1Path,
                Tooltip = "Picture taken from Alosh Bennett https://flic.kr/p/5fKBTq",
                ContextLink = "https://flic.kr/p/5fKBTq",
                Source = "https://flic.kr/p/5fKBTq"
            };
        }
        #endregion
    }
}
