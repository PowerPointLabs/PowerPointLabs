using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportSupertipRibbonId(TextCollection.AddSpotlightTag)]
    class AddSpotlightSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.AddSpotlightButtonSupertip;
        }
    }
}
