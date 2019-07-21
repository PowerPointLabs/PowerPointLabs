using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Extension;
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
using PowerPointLabs.PictureSlidesLab.ViewModel.SliderPropHandler.Factory;
using PowerPointLabs.PictureSlidesLab.Views.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
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

        public ObservableInt SelectedSliderValue { get; set; }

        public ObservableBoolean IsSliderValueChanged { get; set; }

        public ObservableInt SelectedSliderMaximum { get; set; }

        public ObservableInt SelectedSliderTickFrequency { get; set; }

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

        [Import(typeof(SliderPropHandlerFactory))]
        private SliderPropHandlerFactory PropHandlerFactory { get; set; }

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
            Logger.Log("Init PSL View Model begins");
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

            AggregateCatalog catalog = new AggregateCatalog(
                new AssemblyCatalog(Assembly.GetExecutingAssembly()));
            CompositionContainer container = new CompositionContainer(catalog);
            container.ComposeParts(this);

            Logger.Log("Init PSL View Model done");
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
            Logger.Log("ViewModel clean up done");
        }
        #endregion

        #region Stage - Image Selection (Add Image)

        public void RemoveAllImageSelectionListItems()
        {
            ImageSelectionList.Clear();
            ImageSelectionList.Add(CreateChoosePicturesItem());
            Logger.Log("Clear all images done");
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
                Logger.Log("Add local picture begins");
                bool isToSelectPicture = ImageSelectionList.Count == 1;
                foreach (string filename in filenames)
                {
                    VerifyIsProperImage(filename);
                    ImageItem fromFileItem = new ImageItem
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
                if (isToSelectPicture)
                {
                    ImageSelectionListSelectedId.Number = 1;
                }
                Logger.Log("Add local picture done");
            }
            catch (Exception e)
            {
                // not an image or image is corrupted
                View.ShowErrorMessageBox(PictureSlidesLabText.ErrorImageCorrupted);
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
                View.ShowErrorMessageBox(PictureSlidesLabText.ErrorUrlLinkIncorrect);
                Logger.Log("Url link error when add internet image");
                return;
            }
            ImageItem item = new ImageItem
            {
                ImageFile = StoragePath.LoadingImgPath,
                ContextLink = downloadLink,
                Source = downloadLink
            };
            UrlUtil.GetMetaInfo(ref downloadLink, item);
            ImageSelectionList.Add(item);
            IsActiveDownloadProgressRing.Flag = true;

            string imagePath = StoragePath.GetPath("img-"
                + DateTime.Now.GetHashCode() + "-"
                + Guid.NewGuid().ToString().Substring(0, 7));
            ImageDownloader
                .Get(downloadLink, imagePath)
                .After((AutoUpdate.Downloader.AfterDownloadEventDelegate)(() =>
                {
                    try
                    {
                        Logger.Log("Add internet picture begins");
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
                        Logger.Log("Add internet picture ends");
                    }
                    catch (Exception e)
                    {
                        View.ShowErrorMessageBox(PictureSlidesLabText.ErrorImageDownloadCorrupted);
                        ImageSelectionList.Remove(item);
                        Logger.LogException(e, "AddImageSelectionListItem (download)");
                    }
                    finally
                    {
                        IsActiveDownloadProgressRing.Flag = false;
                    }
                }))
                // Case 3: Possibly network timeout
                .OnError((AutoUpdate.Downloader.ErrorEventDelegate)(e =>
                {
                    IsActiveDownloadProgressRing.Flag = false;
                    ImageSelectionList.Remove(item);
                    View.ShowErrorMessageBox(PictureSlidesLabText.ErrorFailedToLoad + e.Message);
                }))
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
        public void UpdatePreviewImages(ImageItem source, Slide contentSlide, float slideWidth, float slideHeight, bool isUpdateSelectedPreviewOnly = false)
        {
            if (View.IsVariationsFlyoutOpen)
            {
                Logger.Log("Generate preview images for variation stage");
                UpdateStylesVariationImagesAfterOpenFlyout(source, contentSlide, slideWidth, slideHeight, isUpdateSelectedPreviewOnly);
                Logger.Log("Generate preview images for variation stage done");
            }
            else
            {
                Logger.Log("Generate preview images for preview style stage");
                UpdateStylesPreviewImages(source, contentSlide, slideWidth, slideHeight);
                Logger.Log("Generate preview images for preview style stage done");
            }
        }

        public void ApplyStyleInPreviewStage(Slide contentSlide, float slideWidth, float slideHeight)
        {
            Logger.Log("Apply style in preview stage begins");
            IList<object> copiedPicture = LoadClipboardPicture();
            try
            {
                StyleOption targetDefaultOptions = OptionsFactory
                    .GetStylesPreviewOption(StylesPreviewListSelectedItem.ImageItem.Tooltip);
                Designer.ApplyStyle(ImageSelectionListSelectedItem.ImageItem, contentSlide,
                    slideWidth, slideHeight, targetDefaultOptions);
                View.ShowSuccessfullyAppliedDialog();
            }
            catch (Exception e)
            {
                View.ShowErrorMessageBox(PictureSlidesLabText.ErrorNoSelectedSlide);
                Logger.LogException(e, "ApplyStyleInPreviewStage");
            }
            SaveClipboardPicture(copiedPicture);
            Logger.Log("Apply style in preview stage done");
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
                View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorFailedToRetrieveInfoFromImage, e);
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
            ImageItem targetStyleItem = StylesPreviewListSelectedItem.ImageItem;
            StylesVariationList.Clear();

            if (!IsAbleToUpdateStylesVariationImages(source, targetStyleItem, contentSlide))
            {
                return;
            }

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
                    VariantsCategory.IndexOf(PictureSlidesLabText.VariantCategoryPicture);
            }
            Logger.Log("Variation open completed");
        }

        /// <summary>
        /// Update styles variation images after its flyout been open
        /// </summary>
        public void UpdateStylesVariationImagesAfterOpenFlyout(ImageItem source, Slide contentSlide, float slideWidth, float slideHeight, bool isUpdateSelectedPreviewOnly = false)
        {
            Logger.Log("Variation is already open, update preview images");
            int selectedId = StylesVariationListSelectedId.Number;
            double scrollOffset = View.GetVariationListBoxScrollOffset();
            ImageItem targetStyleItem = StylesPreviewListSelectedItem.ImageItem;
            if (!isUpdateSelectedPreviewOnly)
            {
                StylesVariationList.Clear();
            }

            if (!IsAbleToUpdateStylesVariationImages(source, targetStyleItem, contentSlide))
            {
                return;
            }

            StylesVariationListSelectedId.Number = selectedId < 0 ? 0 : selectedId;

            if (isUpdateSelectedPreviewOnly)
            {
                UpdateStylesVariationImages(source, contentSlide, slideWidth, slideHeight, selectedId: StylesVariationListSelectedId.Number);
            }
            else
            {
                UpdateStylesVariationImages(source, contentSlide, slideWidth, slideHeight);
            }

            View.SetVariationListBoxScrollOffset(scrollOffset);
        }

        /// <summary>
        /// This method implements the way to guide the user step by step to customize
        /// style
        /// </summary>
        public void UpdateStepByStepStylesVariationImages(ImageItem source, Slide contentSlide, 
            float slideWidth, float slideHeight)
        {
            Logger.Log("Check for step by step preview");
            Logger.Log("current variation list selected id is " + StylesVariationListSelectedId.Number);
            Logger.Log("variants category count is " + VariantsCategory.Count);
            if (StylesVariationListSelectedId.Number < 0
                || VariantsCategory.Count == 0)
            {
                return;
            }

            Logger.Log("Step by step preview begins");
            int targetVariationSelectedIndex = StylesVariationListSelectedId.Number;
            StyleVariant targetVariant = _styleVariants[_previousVariantsCategory][targetVariationSelectedIndex];
            foreach (StyleOption option in _styleOptions)
            {
                targetVariant.Apply(option);
            }
            
            string currentVariantsCategory = CurrentVariantCategory.Text;
            if (currentVariantsCategory != PictureSlidesLabText.VariantCategoryFontColor
                && _previousVariantsCategory != PictureSlidesLabText.VariantCategoryFontColor)
            {
                // apply font color variant,
                // because default styles may contain special font color settings, but not in variants
                StyleVariant fontColorVariant = new StyleVariant(new Dictionary<string, object>
                {
                    {"FontColor", _styleOptions[targetVariationSelectedIndex].FontColor}
                });
                foreach (StyleOption option in _styleOptions)
                {
                    fontColorVariant.Apply(option);
                }
            }

            List<StyleVariant> nextCategoryVariants = _styleVariants[currentVariantsCategory];
            if (currentVariantsCategory == PictureSlidesLabText.VariantCategoryFontFamily)
            {
                bool isFontInVariation = false;
                string currentFontFamily = _styleOptions[targetVariationSelectedIndex].FontFamily;
                foreach (StyleVariant variant in nextCategoryVariants)
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
            if (CurrentVariantCategory.Text == PictureSlidesLabText.VariantCategoryPicture
                && !_isPictureVariationInit)
            {
                _8PicturesInPictureVariation = GetLast8Pictures(targetVariationSelectedIndex);
                _isPictureVariationInit = true;
            }
            // Enter picture variation again
            else if (CurrentVariantCategory.Text == PictureSlidesLabText.VariantCategoryPicture
                     && _isPictureVariationInit)
            {
                bool isPictureSwapped = false;
                for (int i = 0; i < _8PicturesInPictureVariation.Count; i++)
                {
                    // swap the picture to the current selected id in
                    // variation list
                    ImageItem picture = _8PicturesInPictureVariation[i];
                    if ((ImageSelectionListSelectedItem.ImageItem == null 
                        && picture.ImageFile == StoragePath.NoPicturePlaceholderImgPath) || 
                            (ImageSelectionListSelectedItem.ImageItem != null
                            && picture.ImageFile == ImageSelectionListSelectedItem.ImageItem.ImageFile))
                    {
                        ImageItem tempPic = _8PicturesInPictureVariation[targetVariationSelectedIndex];
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
            else if (_previousVariantsCategory == PictureSlidesLabText.VariantCategoryPicture)
            {
                // use the selected picture in the picture variation to preview
                ImageItem targetPicture = _8PicturesInPictureVariation[targetVariationSelectedIndex];
                if (targetPicture.ImageFile != StoragePath.NoPicturePlaceholderImgPath)
                {
                    int indexForTargetPicture = ImageSelectionList.IndexOf(targetPicture);
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
                    View.EnterDefaultPictureMode();
                    source = View.CreateDefaultPictureItem();
                }
            }

            int variantIndexWithoutEffect = -1;
            for (int i = 0; i < nextCategoryVariants.Count; i++)
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
                StyleVariant temp = nextCategoryVariants[variantIndexWithoutEffect];
                nextCategoryVariants[variantIndexWithoutEffect] =
                    nextCategoryVariants[targetVariationSelectedIndex];
                nextCategoryVariants[targetVariationSelectedIndex] = temp;
            }

            for (int i = 0; i < nextCategoryVariants.Count && i < _styleOptions.Count; i++)
            {
                nextCategoryVariants[i].Apply(_styleOptions[i]);
            }

            _previousVariantsCategory = currentVariantsCategory;
            Logger.Log("picture index to select is " + pictureIndexToSelect);
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
            Logger.Log("Step by step preview done");
        }

        public void ApplyStyleInVariationStage(Slide contentSlide, float slideWidth, float slideHeight)
        {
            Logger.Log("Apply style in variation stage begins");
            IList<object> copiedPicture = LoadClipboardPicture();
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
                    ImageItem targetPicture = GetSelectedPictureInPictureVariation(
                        StylesVariationListSelectedId.Number);
                    if (targetPicture.ImageFile != StoragePath.NoPicturePlaceholderImgPath)
                    {
                        int indexForTargetPicture = ImageSelectionList.IndexOf(targetPicture);
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
                View.ShowErrorMessageBox(PictureSlidesLabText.ErrorNoSelectedSlide);
                Logger.LogException(e, "ApplyStyleInVariationStage");
            }
            SaveClipboardPicture(copiedPicture);
            Logger.Log("Apply style in variation stage done");
        }

        #region Picture Variation

        public bool IsInPictureVariation()
        {
            return CurrentVariantCategory != null && CurrentVariantCategory.Text != null
                   && CurrentVariantCategory.Text == PictureSlidesLabText.VariantCategoryPicture;
        }

        public ImageItem GetSelectedPictureInPictureVariation(int pictureIndex)
        {
            try
            {
                return _8PicturesInPictureVariation[pictureIndex];
            }
            catch (Exception e)
            {
                View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorFailedToFetchPictureAspect, e);
                Logger.LogException(e, "GetSelectedPictureInPictureVariation");
                return View.CreateDefaultPictureItem();
            }
        }

        public void UpdateSelectedPictureInPictureVariation()
        {
            Logger.Log("Check for update selected picture in picture aspect");
            Logger.Log("is in picture aspect: " + IsInPictureVariation());
            Logger.Log("variation list selectedId is: " + StylesVariationListSelectedId.Number);
            if (!IsInPictureVariation()
                || StylesVariationListSelectedId.Number == -1)
            {
                return;
            }

            try
            {
                _8PicturesInPictureVariation[StylesVariationListSelectedId.Number]
                    = ImageSelectionListSelectedItem.ImageItem ?? View.CreateDefaultPictureItem();
            }
            catch (Exception e)
            {
                View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorFailedToProcessPictureAspect, e);
                Logger.LogException(e, "UpdateSelectedPictureInPictureVariation");
            }
        }

        public void UpdatePictureInPictureVariationWhenAddedNewOne(ImageItem newPicture)
        {
            Logger.Log("Check for update picture in picture aspect when added new one");
            Logger.Log("is in picture aspect: " + IsInPictureVariation());
            Logger.Log("new pic is null: " + (newPicture == null));
            if (!IsInPictureVariation() || newPicture == null)
            {
                return;
            }

            for (int i = 0; i < _8PicturesInPictureVariation.Count; i++)
            {
                ImageItem imageItem = _8PicturesInPictureVariation[i];
                if (imageItem.ImageFile == StoragePath.NoPicturePlaceholderImgPath)
                {
                    _8PicturesInPictureVariation[i] = newPicture;
                    break;
                }
            }
        }

        public void UpdatePictureInPictureVariationWhenDeleteSome()
        {
            Logger.Log("Check for update picture in picture aspect when deleted some");
            Logger.Log("is in picture aspect: " + IsInPictureVariation());
            if (!IsInPictureVariation())
            {
                return;
            }

            for (int i = 0; i < _8PicturesInPictureVariation.Count; i++)
            {
                ImageItem imageItem = _8PicturesInPictureVariation[i];
                if (ImageSelectionList.IndexOf(imageItem) == -1)
                {
                    _8PicturesInPictureVariation[i] = View.CreateDefaultPictureItem();
                }
            }
        }

        public void RefreshLast8Pictures()
        {
            int selectedIdOfVariationList = Math.Max(StylesVariationListSelectedId.Number, 0);
            _8PicturesInPictureVariation = GetLast8Pictures(selectedIdOfVariationList);
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

        private List<ImageItem> GetLast8Pictures(int selectedIdOfVariationList)
        {
            if (!IsInPictureVariation())
            {
                return new List<ImageItem>();
            }

            try
            {
                IEnumerable<ImageItem> subPictureList = ImageSelectionList.Skip(Math.Max(1, ImageSelectionList.Count - 8));
                List<ImageItem> result = new List<ImageItem>(subPictureList);
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
                    for (int i = 0; i < result.Count; i++)
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
                    int indexToSwap = result.IndexOf(ImageSelectionListSelectedItem.ImageItem);
                    ImageItem tempItem = result[selectedIdOfVariationList];
                    result[selectedIdOfVariationList] = ImageSelectionListSelectedItem.ImageItem;
                    result[indexToSwap] = tempItem;
                }
                return result;
            }
            catch (Exception e)
            {
                View.ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorFailedToGeneratePictureAspect, e);
                Logger.LogException(e, "GetLast8Pictures");
                return new List<ImageItem>();
            }
        }

        private static IList<object> LoadClipboardPicture()
        {
            try
            {
                Logger.Log("Load clipboard begins.");
                List<object> result = PPLClipboard.Instance.LoadClipboardObjects();
                Logger.Log("Load clipboard done.");
                return result;
            }
            catch (Exception e)
            {
                // sometimes Clipboard may fail
                Logger.LogException(e, "LoadClipboardPicture");
                return new List<object>();
            }
        }

        private static void SaveClipboardPicture(IList<object> copiedObjs)
        {
            try
            {
                Logger.Log("Save clipboard begins.");
                foreach (object copiedObj in copiedObjs)
                {
                    if (copiedObj == null)
                    {
                        continue;
                    }

                    if (copiedObj is Image)
                    {
                        Clipboard.SetImage((Image)copiedObj);
                    }
                    else if (copiedObj is StringCollection)
                    {
                        Clipboard.SetFileDropList((StringCollection)copiedObj);
                    }
                    else if (!string.IsNullOrEmpty(copiedObj as string))
                    {
                        Clipboard.SetText((string)copiedObj);
                    }
                }
                Logger.Log("Save clipboard done.");
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
            Logger.Log("UpdateStylesPreviewImages begins");
            int selectedId = StylesPreviewListSelectedId.Number;
            StylesPreviewList.Clear();

            if (!IsAbleToUpdateStylesPreviewImages(source, contentSlide))
            {
                return;
            }

            IList<object> copiedPicture = LoadClipboardPicture();
            try
            {
                List<StyleOption> allStyleOptions = OptionsFactory.GetAllStylesPreviewOptions();
                Logger.Log("Number of styles: " + allStyleOptions.Count);
                foreach (StyleOption stylesPreviewOption in allStyleOptions)
                {
                    PreviewInfo previewInfo = Designer.PreviewApplyStyle(source, 
                        contentSlide, slideWidth, slideHeight, stylesPreviewOption);
                    StylesPreviewList.Add(new ImageItem
                    {
                        ImageFile = previewInfo.PreviewApplyStyleImagePath,
                        Tooltip = stylesPreviewOption.StyleName
                    });
                }
                GC.Collect();
            }
            catch (Exception e)
            {
                View.ShowErrorMessageBox(PictureSlidesLabText.ErrorImageCorrupted, e);
                Logger.LogException(e, "UpdateStylesPreviewImages");
            }
            SaveClipboardPicture(copiedPicture);

            StylesPreviewListSelectedId.Number = selectedId < 0 ? 0 : selectedId;
            Logger.Log("UpdateStylesPreviewImages done");
        }

        private static bool IsAbleToUpdateStylesPreviewImages(ImageItem source, Slide contentSlide)
        {
            Logger.Log("Check for update styles in preview styles stage");
            Logger.Log("source is null: " + (source == null));
            Logger.Log("source is loading img: " + (source != null && source.ImageFile == StoragePath.LoadingImgPath));
            Logger.Log("content slide is null: " + (contentSlide == null));
            return !(source == null
                    || source.ImageFile == StoragePath.LoadingImgPath
                    || contentSlide == null);
        }

        private void InitStylesVariationCategories(List<StyleOption> givenOptions,
            Dictionary<string, List<StyleVariant>> givenVariants, string targetStyle)
        {
            Logger.Log("Init variation stage begins");
            _styleOptions = givenOptions ?? OptionsFactory.GetStylesVariationOptions(targetStyle);
            _styleVariants = givenVariants ?? VariantsFactory.GetVariants(targetStyle);

            VariantsCategory.Clear();
            foreach (string styleVariant in _styleVariants.Keys)
            {
                VariantsCategory.Add(styleVariant);
            }
            CurrentVariantCategoryId.Number = 0;
            _previousVariantsCategory = VariantsCategory[0];

            // default style options (in preview stage)
            StyleOption defaultStyleOptions = OptionsFactory.GetStylesPreviewOption(targetStyle);
            List<StyleVariant> currentVariants = _styleVariants.Values.First();
            int variantIndexWithoutEffect = -1;
            for (int i = 0; i < currentVariants.Count; i++)
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
                StyleVariant tempVariant = currentVariants[variantIndexWithoutEffect];
                currentVariants[variantIndexWithoutEffect] =
                    currentVariants[0];
                currentVariants[0] = tempVariant;
                // swap default style options (in variation stage)
                StyleOption tempStyleOpt = _styleOptions[variantIndexWithoutEffect];
                _styleOptions[variantIndexWithoutEffect] =
                    _styleOptions[0];
                _styleOptions[0] = tempStyleOpt;
            }

            for (int i = 0; i < currentVariants.Count && i < _styleOptions.Count; i++)
            {
                currentVariants[i].Apply(_styleOptions[i]);
            }
            Logger.Log("Init variation stage done");
        }

        private static bool IsAbleToUpdateStylesVariationImages(ImageItem source, ImageItem targetStyleItem, 
            Slide contentSlide)
        {
            Logger.Log("Check for update styles in variation stage");
            Logger.Log("source is null: " + (source == null));
            Logger.Log("source is loading img: " + (source != null && source.ImageFile == StoragePath.LoadingImgPath));
            Logger.Log("target style item is null: " + (targetStyleItem == null));
            Logger.Log("target style item tooltip is null: " + (targetStyleItem != null && targetStyleItem.Tooltip == null));
            Logger.Log("content slide is null: " + (contentSlide == null));
            return !(source == null
                    || source.ImageFile == StoragePath.LoadingImgPath
                    || targetStyleItem == null
                    || targetStyleItem.Tooltip == null
                    || contentSlide == null);
        }

        private void UpdateStylesVariationImages(ImageItem source, Slide contentSlide,
            float slideWidth, float slideHeight, bool isMockPreviewImages = false, int selectedId = -1)
        {
            Logger.Log("UpdateStylesVariationImages begins");
            IList<object> copiedPicture = LoadClipboardPicture();
            try
            {
                if (isMockPreviewImages)
                {
                    Logger.Log("Generate mock images for Picture aspect");
                }
                Logger.Log("Number of styles: " + _styleOptions.Count);
                if (selectedId != -1)
                {
                    StylesVariationList[selectedId] =
                        GenerateImageItem(source, contentSlide, slideWidth, slideHeight, isMockPreviewImages, selectedId);
                }
                else
                {
                    for (int i = 0; i < _styleOptions.Count; i++)
                    {
                        StylesVariationList.Add(
                            GenerateImageItem(source, contentSlide, slideWidth, slideHeight, isMockPreviewImages, i));
                    }
                }
                GC.Collect();
            }
            catch (Exception e)
            {
                View.ShowErrorMessageBox(PictureSlidesLabText.ErrorImageCorrupted, e);
                Logger.LogException(e, "UpdateStylesVariationImages");
            }
            SaveClipboardPicture(copiedPicture);
            Logger.Log("UpdateStylesVariationImages done");
        }

        private ImageItem GenerateImageItem(ImageItem source, Slide contentSlide, float slideWidth, float slideHeight, bool isMockPreviewImages,
            int index)
        {
            StyleOption styleOption = _styleOptions[index];
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
                        ? _8PicturesInPictureVariation[index]
                        : source,
                    contentSlide, slideWidth, slideHeight, styleOption);
            }

            return new ImageItem
            {
                ImageFile = previewInfo.PreviewApplyStyleImagePath,
                Tooltip = styleOption.OptionName
            };
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

        #region Private Lifecycle
        private void InitFontFamilies()
        {
            FontFamilies = new ObservableCollection<string>();
            foreach (System.Windows.Media.FontFamily fontFamily in Fonts.SystemFontFamilies)
            {
                FontFamilies.Add(fontFamily.Source);
            }

            // add font family not in Fonts.SystemFontFamilies
            FontFamilies.Add("Segoe UI Light");
            FontFamilies.Add("Calibri Light");
            FontFamilies.Add("Arial Black");
            FontFamilies.Add("Times New Roman Italic");

            FontFamilies = new ObservableCollection<string>(FontFamilies.OrderBy(i => i));
        }

        private void CleanUnusedPersistentData()
        {
            HashSet<string> imageFilesInUse = new HashSet<string>();
            foreach (ImageItem imageItem in ImageSelectionList)
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
            StylesVariationListSelectedId = new ObservableInt { Number = -1 };
            StylesVariationListSelectedItem = new ObservableImageItem();
            CurrentVariantCategory = new ObservableString();
            CurrentVariantCategoryId = new ObservableInt { Number = -1 };
            VariantsCategory = new ObservableCollection<string>();
            SelectedFontId = new ObservableInt();
            SelectedFontFamily = new ObservableFont();
            SelectedSliderValue = new ObservableInt();
            IsSliderValueChanged = new ObservableBoolean { Flag = false };
            SelectedSliderMaximum = new ObservableInt();
            SelectedSliderTickFrequency = new ObservableInt();

            StylesPreviewList = new ObservableCollection<ImageItem>();
            StylesPreviewListSelectedId = new ObservableInt { Number = -1 };
            StylesPreviewListSelectedItem = new ObservableImageItem();

            ImageSelectionList = new ObservableCollection<ImageItem>();
            ImageSelectionList.Add(CreateChoosePicturesItem());

            Settings = StoragePath.LoadSettings();

            if (StoragePath.IsFirstTimeUsage())
            {
                Logger.Log("First time use PSL");
                ImageSelectionList.Add(CreateSamplePic1Item());
                ImageSelectionList.Add(CreateSamplePic2Item());
            }
            else
            {
                ObservableCollection<ImageItem> loadedImageSelectionList = StoragePath.LoadPictures();
                foreach (ImageItem item in loadedImageSelectionList)
                {
                    if (item.FullSizeImageFile == null && item.BackupFullSizeImageFile != null)
                    {
                        item.FullSizeImageFile = item.BackupFullSizeImageFile;
                    }
                    else if (item.FullSizeImageFile == null && item.BackupFullSizeImageFile == null)
                    {
                        Logger.Log("Corrupted picture found. To be removed");
                        continue;
                    }
                    ImageSelectionList.Add(item);
                }
            }

            ImageSelectionListSelectedId = new ObservableInt { Number = -1 };
            ImageSelectionListSelectedItem = new ObservableImageItem();
            IsActiveDownloadProgressRing = new ObservableBoolean { Flag = false };
        }

        private void InitStorage()
        {
            bool isTempPathInit = Util.TempPath.InitTempFolder();
            bool isStoragePathInit = StoragePath.InitPersistentFolder();
            if (!isTempPathInit || !isStoragePathInit)
            {
                View.ShowErrorMessageBox(PictureSlidesLabText.ErrorFailToInitTempFolder);
                Logger.Log("Failed to init storage");
            }
            Logger.Log("Init storage done");
        }
        #endregion
    }
}
