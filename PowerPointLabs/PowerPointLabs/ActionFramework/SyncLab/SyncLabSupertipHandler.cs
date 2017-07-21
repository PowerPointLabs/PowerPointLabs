using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.SyncLab
{
    [ExportSupertipRibbonId(TextCollection1.SyncLabTag)]
    class SyncLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return SyncLabText.RibbonMenuSupertip;
        }
    }
}
