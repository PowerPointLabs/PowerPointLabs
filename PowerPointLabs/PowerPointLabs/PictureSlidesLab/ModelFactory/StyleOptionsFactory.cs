using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.Linq;
using System.Reflection;

using PowerPointLabs.PictureSlidesLab.Model;
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
        [ImportMany(typeof(IStyleOptions))]
        private IEnumerable<Lazy<IStyleOptions, IStyleOrderMetadata>> ImportedStyleOptions { get; set; }

        public StyleOptionsFactory()
        {
            AggregateCatalog catalog = new AggregateCatalog(
                new AssemblyCatalog(Assembly.GetExecutingAssembly()));
            CompositionContainer container = new CompositionContainer(catalog);
            container.ComposeParts(this);
        }

        /// <summary>
        /// get all styles variation options for variation stage usage
        /// </summary>
        /// <returns></returns>
        public List<List<StyleOption>> GetAllStylesVariationOptions()
        {
            List<List<StyleOption>> options = new List<List<StyleOption>>();
            foreach (IStyleOptions styleOptions in GetAllStyleOptions())
            {
                options.Add(styleOptions.GetOptionsForVariation());
            }
            return options;
        }

        /// <summary>
        /// get all styles preview options for preview stage usage
        /// </summary>
        /// <returns></returns>
        public List<StyleOption> GetAllStylesPreviewOptions()
        {
            List<StyleOption> options = new List<StyleOption>();
            foreach (IStyleOptions styleOptions in GetAllStyleOptions())
            {
                options.Add(styleOptions.GetDefaultOptionForPreview());
            }
            return options;
        }

        public StyleOption GetStylesPreviewOption(string targetStyle)
        {
            List<StyleOption> options = GetAllStylesPreviewOptions();
            foreach (StyleOption option in options)
            {
                if (option.StyleName == targetStyle)
                {
                    return option;
                }
            }
            return options[0];
        }

        public List<StyleOption> GetStylesVariationOptions(string targetStyle)
        {
            List<List<StyleOption>> allStylesVariationOptions = GetAllStylesVariationOptions();
            foreach (List<StyleOption> stylesVariationOptions in allStylesVariationOptions)
            {
                if (stylesVariationOptions[0].StyleName == targetStyle)
                {
                    return stylesVariationOptions;
                }
            }
            return allStylesVariationOptions[0];
        }

        public IEnumerable<IStyleOptions> GetAllStyleOptions()
        {
            return ImportedStyleOptions
                .OrderBy(options => options.Metadata.StyleOrder)
                .Select(options => options.Value);
        }
    }
}
