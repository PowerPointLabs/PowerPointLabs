using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.SyncLab
{
    [ExportSupertipRibbonId(SyncLabText.PaneTag)]
    class SyncLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return SyncLabText.RibbonMenuSupertip;
        }
    }
}
