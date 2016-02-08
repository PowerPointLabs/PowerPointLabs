using System.Collections.Generic;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Variants;

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
        public static Dictionary<string, List<StyleVariant>> GetVariants(string targetStyle)
        {
            switch (targetStyle)
            {
                case TextCollection.PictureSlidesLabText.StyleNameDirectText:
                    return new DirectTextStyleVariants().GetVariantsForStyle();
                case TextCollection.PictureSlidesLabText.StyleNameBlur:
                    return new BlurStyleVariants().GetVariantsForStyle();
                case TextCollection.PictureSlidesLabText.StyleNameTextBox:
                    return new TextBoxStyleVariants().GetVariantsForStyle();
                case TextCollection.PictureSlidesLabText.StyleNameBanner:
                    return new BannerStyleVariants().GetVariantsForStyle();
                case TextCollection.PictureSlidesLabText.StyleNameSpecialEffect:
                    return new SpecialEffectStyleVariants().GetVariantsForStyle();
                case TextCollection.PictureSlidesLabText.StyleNameOverlay:
                    return new OverlayStyleVariants().GetVariantsForStyle();
                case TextCollection.PictureSlidesLabText.StyleNameOutline:
                    return new OutlineStyleVariants().GetVariantsForStyle();
                case TextCollection.PictureSlidesLabText.StyleNameFrame:
                    return new FrameStyleVariants().GetVariantsForStyle();
                case TextCollection.PictureSlidesLabText.StyleNameCircle:
                    return new CircleStyleVariants().GetVariantsForStyle();
                case TextCollection.PictureSlidesLabText.StyleNameTriangle:
                    return new TriangleStyleVariants().GetVariantsForStyle();
                default:
                    return new Dictionary<string, List<StyleVariant>>();
            }
        }
    }
}
