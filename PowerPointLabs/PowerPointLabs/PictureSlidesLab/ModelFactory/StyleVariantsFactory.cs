using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.Linq;
using System.Reflection;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Variants.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory
{
    /// <summary>
    /// StyleVariantsFactory constructs the style variants based on given target style name.
    /// 
    /// To support new style variants,
    /// firstly create any new variant worker (a subclass of IVariantWorker that defines the
    /// style variation), and then create the new style variants (a subclass of BaseStyleVariants)
    /// that specifies which variant workers are needed for this style variants.
    /// </summary>
    public class StyleVariantsFactory
    {

        [ImportMany(typeof(IStyleVariants))]
        private IEnumerable<Lazy<IStyleVariants>> ImportedStyleVariants { get; set; }

        public StyleVariantsFactory()
        {
            AggregateCatalog catalog = new AggregateCatalog(
                new AssemblyCatalog(Assembly.GetExecutingAssembly()));
            CompositionContainer container = new CompositionContainer(catalog);
            container.ComposeParts(this);
        }

        public Dictionary<string, List<StyleVariant>> GetVariants(string targetStyle)
        {
            foreach (IStyleVariants styleVariants in GetAllStyleVariants())
            {
                if (styleVariants.GetStyleName() == targetStyle)
                {
                    return styleVariants.GetVariantsForStyle();
                }
            }
            return new Dictionary<string, List<StyleVariant>>();
        }

        public IEnumerable<IStyleVariants> GetAllStyleVariants()
        {
            return ImportedStyleVariants.Select(variants => variants.Value);
        }
    }
}
