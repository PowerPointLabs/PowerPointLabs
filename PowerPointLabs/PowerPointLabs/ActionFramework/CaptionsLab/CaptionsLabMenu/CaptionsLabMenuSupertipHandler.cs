using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.CaptionsLab
{
    [ExportSupertipRibbonId(CaptionsLabText.RibbonMenuId)]
    class CaptionsLabMenuSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return CaptionsLabText.RibbonMenuSupertip;
        }
    }
}
