using System.Windows.Media;
using Microsoft.Office.Interop.PowerPoint;
using Color = System.Drawing.Color;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PictureSlidesLab.ViewModel
{
    public partial class PictureSlidesLabWindowViewModel
    {
        ///////////////////////////////////////////////////////////////
        // Implemented variation stage controls' binding in ViewModel
        ///////////////////////////////////////////////////////////////

        // TODO add new variation stage control Slide to adjust brightness/blurriness/transparency

        #region Binding funcs for color panel
        public void BindStyleToColorPanel()
        {
            if (!IsAbleToBindProperty()) return;

            var styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            var currentCategory = CurrentVariantCategory.Text;
            var bc = new BrushConverter();

            if (currentCategory.Contains(TextCollection.PictureSlidesLabText.ColorHasEffect))
            {
                var propName = GetPropertyName(currentCategory);
                var type = styleOption.GetType();
                var prop = type.GetProperty(propName);
                var optValue = prop.GetValue(styleOption, null);
                if (!string.IsNullOrEmpty(optValue as string))
                {
                    View.SetVariantsColorPanelBackground((Brush) bc.ConvertFrom(optValue));
                }
            }
        }

        public void BindSelectedColor(Color color, Slide contentSlide, float slideWidth, float slideHeight)
        {
            BindColorToStyle(color);
            BindColorToVariant(color);
            if (View.IsDisplayDefaultPicture())
            {
                UpdatePreviewImages(
                    View.CreateDefaultPictureItem(),
                    contentSlide,
                    slideWidth,
                    slideHeight,
                    isUpdateSelectedPreviewOnly: true);
                BindStyleToColorPanel();
            }
            else
            {
                UpdatePreviewImages(
                    ImageSelectionListSelectedItem.ImageItem ??
                    View.CreateDefaultPictureItem(),
                    contentSlide,
                    slideWidth,
                    slideHeight,
                    isUpdateSelectedPreviewOnly: true);
            }
        }
        #endregion

        #region Binding funcs for font panel
        public void BindStyleToFontPanel()
        {
            if (!IsAbleToBindProperty()) return;

            var styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            var styleFontFamily = styleOption.GetFontFamily();
            var targetIndex = -1;
            for (var i = 0; i < FontFamilies.Count; i++)
            {
                if (styleFontFamily == FontFamilies[i])
                {
                    targetIndex = i;
                    break;
                }
            }
            SelectedFontId.Number = targetIndex;
        }

        public void BindSelectedFont(Slide contentSlide, float slideWidth, float slideHeight)
        {
            BindFontToStyle(SelectedFontFamily.Font.Source);
            BindFontToVariant(SelectedFontFamily.Font.Source);
            UpdatePreviewImages(
                ImageSelectionListSelectedItem.ImageItem ??
                View.CreateDefaultPictureItem(),
                contentSlide,
                slideWidth,
                slideHeight,
                isUpdateSelectedPreviewOnly: true);
        }
        #endregion

        #region Binding funcs for slider
        public void BindStyleToSlider()
        {
            if (!IsAbleToBindProperty()) return;

            var styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            var currentCategory = CurrentVariantCategory.Text;
            var propName = GetPropertyName(currentCategory);
            var propHandler = PropHandlerFactory.GetSliderPropHandler(propName);
            var sliderProperties = propHandler.GetSliderProperties(styleOption);
            SelectedSliderValue.Number = sliderProperties.Value;
            SelectedSliderMaximum.Number = sliderProperties.Maximum;
            SelectedSliderTickFrequency.Number = sliderProperties.TickFrequency;
        }

        public void BindSelectedSliderValue(Slide contentSlide, float slideWidth, float slideHeight)
        {
            BindSliderValueToStyle(SelectedSliderValue.Number);
            BindSliderValueToVariant(SelectedSliderValue.Number);
            if (View.IsDisplayDefaultPicture())
            {
                UpdatePreviewImages(
                    View.CreateDefaultPictureItem(),
                    contentSlide,
                    slideWidth,
                    slideHeight,
                    isUpdateSelectedPreviewOnly: true);
                BindStyleToSlider();
            }
            else
            {
                UpdatePreviewImages(
                    ImageSelectionListSelectedItem.ImageItem ??
                    View.CreateDefaultPictureItem(),
                    contentSlide,
                    slideWidth,
                    slideHeight,
                    isUpdateSelectedPreviewOnly: true);
            }
        }
        #endregion

        #region Helper funcs

        private void BindColorToStyle(Color color)
        {
            if (!IsAbleToBindProperty()) return;

            var styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            var currentCategory = CurrentVariantCategory.Text;
            var targetColor = StringUtil.GetHexValue(color);

            if (currentCategory.Contains(TextCollection.PictureSlidesLabText.ColorHasEffect))
            {
                styleOption.OptionName = "Customized";
                var propName = GetPropertyName(currentCategory);
                var type = styleOption.GetType();
                var prop = type.GetProperty(propName);
                prop.SetValue(styleOption, targetColor, null);
            }
        }

        private void BindColorToVariant(Color color)
        {
            if (!IsAbleToBindProperty()) return;

            var currentCategory = CurrentVariantCategory.Text;
            var styleVariant = _styleVariants[currentCategory][StylesVariationListSelectedId.Number];

            if (currentCategory.Contains(TextCollection.PictureSlidesLabText.ColorHasEffect))
            {
                styleVariant.Set("OptionName", "Customized");
                styleVariant.Set(GetPropertyName(currentCategory), StringUtil.GetHexValue(color));
            }
        }

        private void BindFontToStyle(string font)
        {
            if (!IsAbleToBindProperty()) return;

            var styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            var currentCategory = CurrentVariantCategory.Text;

            if (currentCategory == TextCollection.PictureSlidesLabText.VariantCategoryFontFamily)
            {
                styleOption.OptionName = "Customized";
                styleOption.FontFamily = font;
            }
        }

        private void BindFontToVariant(string font)
        {
            if (!IsAbleToBindProperty()) return;

            var currentCategory = CurrentVariantCategory.Text;
            var styleVariant = _styleVariants[currentCategory][StylesVariationListSelectedId.Number];

            if (currentCategory == TextCollection.PictureSlidesLabText.VariantCategoryFontFamily)
            {
                styleVariant.Set("OptionName", "Customized");
                styleVariant.Set("FontFamily", font);
            }
        }

        private void BindSliderValueToStyle(int value)
        {
            if (!IsAbleToBindProperty()) return;

            var styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            var currentCategory = CurrentVariantCategory.Text;
            var propName = GetPropertyName(currentCategory);
            PropHandlerFactory.GetSliderPropHandler(propName).BindStyleOption(styleOption, value);
        }

        private void BindSliderValueToVariant(int value)
        {
            if (!IsAbleToBindProperty()) return;

            var currentCategory = CurrentVariantCategory.Text;
            var styleVariant = _styleVariants[currentCategory][StylesVariationListSelectedId.Number];
            var propName = GetPropertyName(currentCategory);
            PropHandlerFactory.GetSliderPropHandler(propName).BindStyleVariant(styleVariant, value);
        }

        private bool IsAbleToBindProperty()
        {
            return !(StylesVariationListSelectedId.Number < 0
                     || VariantsCategory.Count == 0);
        }

        private string GetPropertyName(string categoryName)
        {
            var propName = categoryName.Replace(" ", string.Empty);

            var styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            if ((styleOption.IsUseFrostedGlassBannerStyle && categoryName.Contains(TextCollection.PictureSlidesLabText.BannerHasEffect))
                || (styleOption.IsUseFrostedGlassTextBoxStyle && categoryName.Contains(TextCollection.PictureSlidesLabText.TextBoxHasEffect)))
            {
                propName = propName.Insert(0, "FrostedGlass");
            }

            return propName;
        }
        #endregion
    }
}
