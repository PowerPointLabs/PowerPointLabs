using System.Linq;
using System.Windows.Media;
using PowerPointLabs.Models;
using Color = System.Drawing.Color;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PictureSlidesLab.ViewModel
{
    partial class PictureSlidesLabWindowViewModel
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

            if (currentCategory.Contains("Color"))
            {
                var propName = GetPropertyName(currentCategory);
                var type = styleOption.GetType();
                var prop = type.GetProperty(propName);
                var optValue = prop.GetValue(styleOption, null);
                View.SetVariantsColorPanelBackground((Brush)bc.ConvertFrom(optValue));
            }
        }

        public void BindSelectedColor(Color color)
        {
            BindColorToStyle(color);
            BindColorToVariant(color);
            // it will auto update preview images, because PictureSlidesLabWindow will re-activate
        }
        #endregion

        #region Binding funcs for font panel
        public void BindStyleToFontPanel()
        {
            if (!IsAbleToBindProperty()) return;

            var styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            var styleFontFamily = styleOption.GetFontFamily();
            var targetIndex = -1;
            for (var i = 0; i < Fonts.SystemFontFamilies.Count; i++)
            {
                if (styleFontFamily == Fonts.SystemFontFamilies.ElementAt(i).Source)
                {
                    targetIndex = i;
                    break;
                }
            }
            SelectedFontId.Number = targetIndex;
        }

        public void BindSelectedFont()
        {
            BindFontToStyle(SelectedFontFamily.Font.Source);
            BindFontToVariant(SelectedFontFamily.Font.Source);
            UpdatePreviewImages(
                PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide(),
                PowerPointPresentation.Current.SlideWidth,
                PowerPointPresentation.Current.SlideHeight);
        }
        #endregion

        #region Helper funcs

        private void BindColorToStyle(Color color)
        {
            if (!IsAbleToBindProperty()) return;

            var styleOption = _styleOptions[StylesVariationListSelectedId.Number];
            var currentCategory = CurrentVariantCategory.Text;
            var targetColor = StringUtil.GetHexValue(color);

            if (currentCategory.Contains("Color"))
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

            if (currentCategory.Contains("Color"))
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
