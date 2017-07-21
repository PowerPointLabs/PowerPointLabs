using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.SyncLab
{
    [ExportSupertipRibbonId(TextCollection1.SyncLabTag)]
    class SyncLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection1.SyncLabMenuSupertip;
        }
    }
}
