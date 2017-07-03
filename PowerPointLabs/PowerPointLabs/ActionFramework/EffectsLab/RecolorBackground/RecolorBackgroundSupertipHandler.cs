using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportSupertipRibbonId(TextCollection.RecolorBackgroundMenuId)]
    class RecolorBackgroundSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.EffectsLabRecolorBackgroundSupertip;
        }
    }
}
