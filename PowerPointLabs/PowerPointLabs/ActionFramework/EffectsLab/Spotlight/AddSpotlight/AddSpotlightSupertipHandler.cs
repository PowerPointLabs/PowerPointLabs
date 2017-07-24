using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportSupertipRibbonId(EffectsLabText.AddSpotlightTag)]
    class AddSpotlightSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return EffectsLabText.AddSpotlightButtonSupertip;
        }
    }
}
