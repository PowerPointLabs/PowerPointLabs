using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.CaptionsLab
{
    [ExportSupertipRibbonId(TextCollection.MakeTransparentTag)]
    class AddCaptionsSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.EffectsLabMakeTransparentSupertip;
        }
    }
}
