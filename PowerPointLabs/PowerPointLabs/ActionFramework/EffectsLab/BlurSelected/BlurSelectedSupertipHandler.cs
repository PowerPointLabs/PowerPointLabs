using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportSupertipRibbonId(TextCollection.BlurSelectedMenuId)]
    class BlurSelectedSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.EffectsLabBlurSelectedSupertip;
        }
    }
}
