using System.Collections.Generic;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory
{
    /// <summary>
    /// in order to ensure continuity in the customisation stage,
    /// style option provided from this factory should have corresponding values specified 
    /// in StyleVariantsFactory. e.g., an option generated from this factory has overlay 
    /// transparency of 35, then in order to swap (ensure continuity), it should have a 
    /// variant of overlay transparency of 35. Otherwise it cannot swap and so lose continuity, 
    /// because variants don't match any values in the style option.
    /// </summary>
    public class StyleOptionsFactory
    {
        /// <summary>
        /// Add new style options here
        /// </summary>
        /// <returns></returns>
        private static List<IStyleOptions> GetAllStyleOptions()
        {
            return new List<IStyleOptions>
            {
                new DirectTextStyleOptions(),
                new BlurStyleOptions(),
                new TextBoxStyleOptions(),
                new BannerStyleOptions(),
                new SpecialEffectStyleOptions(),
                new OverlayStyleOptions(),
                new OutlineStyleOptions(),
                new FrameStyleOptions(),
                new CircleStyleOptions(),
                new TriangleStyleOptions()
            };
        }

        /// <summary>
        /// get all styles variation options for variation stage usage
        /// </summary>
        /// <returns></returns>
        public static List<List<StyleOption>> GetAllStylesVariationOptions()
        {
            var options = new List<List<StyleOption>>();
            foreach (var styleOptions in GetAllStyleOptions())
            {
                options.Add(styleOptions.GetOptionsForVariation());
            }
            return options;
        }

        /// <summary>
        /// get all styles preview options for preview stage usage
        /// </summary>
        /// <returns></returns>
        public static List<StyleOption> GetAllStylesPreviewOptions()
        {
            var options = new List<StyleOption>();
            foreach (var styleOptions in GetAllStyleOptions())
            {
                options.Add(styleOptions.GetDefaultOptionForPreview());
            }
            return options;
        }

        public static StyleOption GetStylesPreviewOption(string targetStyle)
        {
            var options = GetAllStylesPreviewOptions();
            foreach (var option in options)
            {
                if (option.StyleName == targetStyle)
                {
                    return option;
                }
            }
            return options[0];
        }

        public static List<StyleOption> GetStylesVariationOptions(string targetStyle)
        {
            var allStylesVariationOptions = GetAllStylesVariationOptions();
            foreach (var stylesVariationOptions in allStylesVariationOptions)
            {
                if (stylesVariationOptions[0].StyleName == targetStyle)
                {
                    return stylesVariationOptions;
                }
            }
            return allStylesVariationOptions[0];
        }
    }
}
