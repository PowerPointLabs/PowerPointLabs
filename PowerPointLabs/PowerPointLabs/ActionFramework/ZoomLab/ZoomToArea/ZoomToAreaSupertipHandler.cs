using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ZoomLab
{
    [ExportSupertipRibbonId(ZoomLabText.ZoomToAreaTag)]
    class ZoomToAreaSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return ZoomLabText.ZoomToAreaButtonSupertip;
        }
    }
}
