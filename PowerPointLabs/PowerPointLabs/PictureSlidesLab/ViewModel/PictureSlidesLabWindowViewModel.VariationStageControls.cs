using System.Windows.Media;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
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
                View.SetVariantsColorPanelBackground((Brush)bc.ConvertFrom(optValue));
            }
        }

        public void BindSelectedColor(Color color, Slide contentSlide, float slideWidth, float slideHeight)
        {
            BindColorToStyle(color);
            BindColorToVariant(color);
            if (View.IsDisplayDefaultPicture())
            {
                View.EnableUpdatingPreviewImages();
                UpdatePreviewImages(
                    View.CreateDefaultPictureItem(),
                    contentSlide,
                    slideWidth,
                    slideHeight);
                View.DisableUpdatingPreviewImages();
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
