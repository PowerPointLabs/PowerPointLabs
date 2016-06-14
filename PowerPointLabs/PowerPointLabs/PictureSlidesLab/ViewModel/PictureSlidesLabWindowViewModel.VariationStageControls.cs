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
                    slideHeight);
                BindStyleToColorPanel();
            }
            else
            {
                UpdatePreviewImages(
                    ImageSelectionListSelectedItem.ImageItem ??
                    View.CreateDefaultPictureItem(),
                    contentSlide,
                    slideWidth,
                    slideHeight);
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
                slideHeight);
        }
        #endregion

        #region Binding funcs for slider
        public void BindStyleToSlider()
        {
            if (!IsAbleToBindProperty()) return;

            var styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            var currentCategory = CurrentVariantCategory.Text;
            var type = styleOption.GetType();

            if (currentCategory == TextCollection.PictureSlidesLabText.VariantCategoryBlurriness)
            {
                var prop = type.GetProperty("BlurDegree");
                var optValue = (int)prop.GetValue(styleOption, null);
                SelectedSliderValue.Number = (optValue - 50) * 2;
                SelectedSliderMaximum.Number = 100;
                SelectedSliderTickFrequency.Number = 2;
            }
            if (currentCategory == TextCollection.PictureSlidesLabText.VariantCategoryBrightness)
            {
                var prop = type.GetProperty("OverlayColor");
                var colorValue = (string)prop.GetValue(styleOption, null);
                prop = type.GetProperty("Transparency");
                var optValue = (int)prop.GetValue(styleOption, null);

                if (colorValue == "#FFFFFF")
                {
                    SelectedSliderValue.Number = 200 - optValue;
                }
                else
                {
                    SelectedSliderValue.Number = optValue;
                }
                
                SelectedSliderMaximum.Number = 200;
                SelectedSliderTickFrequency.Number = 1;
            }
            else if (currentCategory.Contains(TextCollection.PictureSlidesLabText.TransparencyHasEffect))
            {
                var propName = GetPropertyName(currentCategory);
                var prop = type.GetProperty(propName);
                var optValue = (int)prop.GetValue(styleOption, null);
                SelectedSliderValue.Number = optValue;
                SelectedSliderMaximum.Number = 100;
                SelectedSliderTickFrequency.Number = 1;
            }

            SelectedSliderToolTip.Text = SelectedSliderValue.Number + "%";
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
                    slideHeight);
                BindStyleToSlider();
            }
            else
            {
                UpdatePreviewImages(
                    ImageSelectionListSelectedItem.ImageItem ??
                    View.CreateDefaultPictureItem(),
                    contentSlide,
                    slideWidth,
                    slideHeight);
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
            var type = styleOption.GetType();

            if (currentCategory == TextCollection.PictureSlidesLabText.VariantCategoryBlurriness)
            {
                styleOption.OptionName = "Customized";
                var prop = type.GetProperty("IsUseBlurStyle");
                prop.SetValue(styleOption, true, null);
                prop = type.GetProperty("BlurDegree");
                prop.SetValue(styleOption, 50 + (value / 2), null);
            }
            else if (currentCategory == TextCollection.PictureSlidesLabText.VariantCategoryBrightness)
            {
                styleOption.OptionName = "Customized";
                var prop = type.GetProperty("IsUseOverlayStyle");
                prop.SetValue(styleOption, true, null);

                if (value > 100)
                {
                    prop = type.GetProperty("OverlayColor");
                    prop.SetValue(styleOption, "#FFFFFF", null);
                    prop = type.GetProperty("Transparency");
                    prop.SetValue(styleOption, 200 - value, null);
                }
                else
                {
                    prop = type.GetProperty("OverlayColor");
                    prop.SetValue(styleOption, "#000000", null);
                    prop = type.GetProperty("Transparency");
                    prop.SetValue(styleOption, value, null);
                }
            }
            else if (currentCategory.Contains(TextCollection.PictureSlidesLabText.TransparencyHasEffect))
            {
                styleOption.OptionName = "Customized";
                var propName = GetPropertyName(currentCategory);
                var prop = type.GetProperty(propName);
                prop.SetValue(styleOption, value, null);
            }
        }

        private void BindSliderValueToVariant(int value)
        {
            if (!IsAbleToBindProperty()) return;

            var currentCategory = CurrentVariantCategory.Text;
            var styleVariant = _styleVariants[currentCategory][StylesVariationListSelectedId.Number];

            if (currentCategory == TextCollection.PictureSlidesLabText.VariantCategoryBlurriness)
            {
                styleVariant.Set("OptionName", "Customized");
                styleVariant.Set("IsUseBlurStyle", true);
                styleVariant.Set("BlurDegree", 50 + (value / 2));
            }
            else if (currentCategory == TextCollection.PictureSlidesLabText.VariantCategoryBrightness)
            {
                styleVariant.Set("OptionName", "Customized");
                styleVariant.Set("IsUseOverlayStyle", true);

                if (value > 100)
                {
                    styleVariant.Set("OverlayColor", "#FFFFFF");
                    styleVariant.Set("Transparency", 200 - value);
                }
                else
                {
                    styleVariant.Set("OverlayColor", "#000000");
                    styleVariant.Set("Transparency", value);
                }
            }
            else if (currentCategory.Contains(TextCollection.PictureSlidesLabText.TransparencyHasEffect))
            {
                styleVariant.Set("OptionName", "Customized");
                styleVariant.Set(GetPropertyName(currentCategory), value);
            }
        }

        private bool IsAbleToBindProperty()
        {
            return !(StylesVariationListSelectedId.Number < 0
                     || VariantsCategory.Count == 0);
        }

        private string GetPropertyName(string categoryName)
        {
            return categoryName.Replace(" ", string.Empty);
        }
        #endregion
    }
}
