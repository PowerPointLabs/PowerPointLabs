using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.CheckBoxAction
{
    [ExportCheckBoxActionRibbonId(TextCollection.EffectsLabBlurrinessTag)]
    class EffectsLabBlurrinessCheckBoxActionHandler : CheckBoxActionHandler
    {
        protected override void ExecuteCheckBoxAction(string ribbonId, bool pressed)
        {
            var feature = ribbonId.Substring(0, ribbonId.IndexOf(TextCollection.DynamicMenuCheckBoxId));

            switch (feature)
            {
                case TextCollection.EffectsLabBlurrinessFeatureSelected:
                    EffectsLab.EffectsLabBlurSelected.IsTintSelected = pressed;
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureRemainder:
                    EffectsLab.EffectsLabBlurSelected.IsTintRemainder = pressed;
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureBackground:
                    EffectsLab.EffectsLabBlurSelected.IsTintBackground = pressed;
                    break;
                default:
                    throw new System.Exception("Invalid feature");
            }
        }
    }
}
