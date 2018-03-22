using System.Windows.Media;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

using Color = System.Drawing.Color;

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
            if (!IsAbleToBindProperty())
            {
                return;
            }

            Model.StyleOption styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            string currentCategory = CurrentVariantCategory.Text;
            BrushConverter bc = new BrushConverter();

            if (currentCategory.Contains(PictureSlidesLabText.ColorHasEffect))
            {
                string propName = GetPropertyName(currentCategory);
                System.Type type = styleOption.GetType();
                System.Reflection.PropertyInfo prop = type.GetProperty(propName);
                object optValue = prop.GetValue(styleOption, null);
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
            if (!IsAbleToBindProperty())
            {
                return;
            }

            Model.StyleOption styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            string styleFontFamily = styleOption.GetFontFamily();
            int targetIndex = -1;
            for (int i = 0; i < FontFamilies.Count; i++)
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
            if (!IsAbleToBindProperty())
            {
                return;
            }

            Model.StyleOption styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            string currentCategory = CurrentVariantCategory.Text;
            string propName = GetPropertyName(currentCategory);
            SliderPropHandler.Interface.ISliderPropHandler propHandler = PropHandlerFactory.GetSliderPropHandler(propName);
            SliderPropHandler.Factory.SliderPropHandlerFactory.SliderProperties sliderProperties = propHandler.GetSliderProperties(styleOption);
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
            if (!IsAbleToBindProperty())
            {
                return;
            }

            Model.StyleOption styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            string currentCategory = CurrentVariantCategory.Text;
            string targetColor = StringUtil.GetHexValue(color);

            if (currentCategory.Contains(PictureSlidesLabText.ColorHasEffect))
            {
                styleOption.OptionName = "Customized";
                string propName = GetPropertyName(currentCategory);
                System.Type type = styleOption.GetType();
                System.Reflection.PropertyInfo prop = type.GetProperty(propName);
                prop.SetValue(styleOption, targetColor, null);
            }
        }

        private void BindColorToVariant(Color color)
        {
            if (!IsAbleToBindProperty())
            {
                return;
            }

            string currentCategory = CurrentVariantCategory.Text;
            Model.StyleVariant styleVariant = _styleVariants[currentCategory][StylesVariationListSelectedId.Number];

            if (currentCategory.Contains(PictureSlidesLabText.ColorHasEffect))
            {
                styleVariant.Set("OptionName", "Customized");
                styleVariant.Set(GetPropertyName(currentCategory), StringUtil.GetHexValue(color));
            }
        }

        private void BindFontToStyle(string font)
        {
            if (!IsAbleToBindProperty())
            {
                return;
            }

            Model.StyleOption styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            string currentCategory = CurrentVariantCategory.Text;

            if (currentCategory == PictureSlidesLabText.VariantCategoryFontFamily)
            {
                styleOption.OptionName = "Customized";
                styleOption.FontFamily = font;
            }
        }

        private void BindFontToVariant(string font)
        {
            if (!IsAbleToBindProperty())
            {
                return;
            }

            string currentCategory = CurrentVariantCategory.Text;
            Model.StyleVariant styleVariant = _styleVariants[currentCategory][StylesVariationListSelectedId.Number];

            if (currentCategory == PictureSlidesLabText.VariantCategoryFontFamily)
            {
                styleVariant.Set("OptionName", "Customized");
                styleVariant.Set("FontFamily", font);
            }
        }

        private void BindSliderValueToStyle(int value)
        {
            if (!IsAbleToBindProperty())
            {
                return;
            }

            Model.StyleOption styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            string currentCategory = CurrentVariantCategory.Text;
            string propName = GetPropertyName(currentCategory);
            PropHandlerFactory.GetSliderPropHandler(propName).BindStyleOption(styleOption, value);
        }

        private void BindSliderValueToVariant(int value)
        {
            if (!IsAbleToBindProperty())
            {
                return;
            }

            string currentCategory = CurrentVariantCategory.Text;
            Model.StyleVariant styleVariant = _styleVariants[currentCategory][StylesVariationListSelectedId.Number];
            string propName = GetPropertyName(currentCategory);
            PropHandlerFactory.GetSliderPropHandler(propName).BindStyleVariant(styleVariant, value);
        }

        private bool IsAbleToBindProperty()
        {
            return !(StylesVariationListSelectedId.Number < 0
                     || VariantsCategory.Count == 0);
        }

        private string GetPropertyName(string categoryName)
        {
            string propName = categoryName.Replace(" ", string.Empty);

            Model.StyleOption styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            if ((styleOption.IsUseFrostedGlassBannerStyle && categoryName.Contains(PictureSlidesLabText.BannerHasEffect))
                || (styleOption.IsUseFrostedGlassTextBoxStyle && categoryName.Contains(PictureSlidesLabText.TextBoxHasEffect)))
            {
                propName = propName.Insert(0, "FrostedGlass");
            }

            return propName;
        }
        #endregion
    }
}
