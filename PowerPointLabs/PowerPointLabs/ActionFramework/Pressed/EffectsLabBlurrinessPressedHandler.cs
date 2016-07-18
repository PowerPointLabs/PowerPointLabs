using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Pressed
{
    [ExportPressedRibbonId(TextCollection.EffectsLabBlurrinessTag)]
    class EffectsLabBlurrinessPressedHandler : PressedHandler
    {
        protected override bool GetPressed(string ribbonId, string ribbonTag)
        {
            return EffectsLab.EffectsLabBlurSelected.HasOverlay;
        }
    }
}
