using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportSupertipRibbonId(EffectsLabText.RibbonMenuId)]
    class EffectsLabMenuSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return EffectsLabText.RibbonMenuSupertip;
        }
    }
}
