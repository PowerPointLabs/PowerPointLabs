using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ResizeLab
{
    [ExportSupertipRibbonId(ResizeLabText.PaneTag)]
    class ResizeLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return ResizeLabText.RibbonMenuSupertip;
        }
    }
}
