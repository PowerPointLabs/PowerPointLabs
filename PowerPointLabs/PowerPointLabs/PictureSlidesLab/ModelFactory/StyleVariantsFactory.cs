using System.Collections.Generic;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Variants;
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
        /// <summary>
        /// Add new style variants here
        /// </summary>
        /// <returns></returns>
        private static List<IStyleVariants> GetAllStyleVariants()
        {
            return new List<IStyleVariants>
            {
                new DirectTextStyleVariants(),
                new BlurStyleVariants(),
                new TextBoxStyleVariants(),
                new BannerStyleVariants(),
                new SpecialEffectStyleVariants(),
                new OverlayStyleVariants(),
                new OutlineStyleVariants(),
                new FrameStyleVariants(),
                new CircleStyleVariants(),
                new TriangleStyleVariants()
            };
        } 

        public static Dictionary<string, List<StyleVariant>> GetVariants(string targetStyle)
        {
            foreach (var styleVariants in GetAllStyleVariants())
            {
                if (styleVariants.GetStyleName() == targetStyle)
                {
                    return styleVariants.GetVariantsForStyle();
                }
            }
            return new Dictionary<string, List<StyleVariant>>();
        }
    }
}
