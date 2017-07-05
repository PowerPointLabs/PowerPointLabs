using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportSupertipRibbonId(TextCollection.RecolorRemainderMenuId)]
    class RecolorRemainderSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.EffectsLabRecolorRemainderSupertip;
        }
    }
}
