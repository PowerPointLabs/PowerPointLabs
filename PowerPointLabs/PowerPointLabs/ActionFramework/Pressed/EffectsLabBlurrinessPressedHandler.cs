using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Pressed
{
    [ExportPressedRibbonId(TextCollection.EffectsLabBlurrinessTag)]
    class EffectsLabBlurrinessPressedHandler : PressedHandler
    {
        protected override bool GetPressed(string ribbonId)
        {
            var feature = ribbonId.Substring(0, ribbonId.IndexOf(TextCollection.DynamicMenuCheckBoxId));

            switch (feature)
            {
                case TextCollection.EffectsLabBlurrinessFeatureSelected:
                    return EffectsLab.EffectsLabBlurSelected.IsTintSelected;
                case TextCollection.EffectsLabBlurrinessFeatureRemainder:
                    return EffectsLab.EffectsLabBlurSelected.IsTintRemainder;
                case TextCollection.EffectsLabBlurrinessFeatureBackground:
                    return EffectsLab.EffectsLabBlurSelected.IsTintBackground;
                default:
                    throw new System.Exception("Invalid feature");
            }
        }
    }
}
