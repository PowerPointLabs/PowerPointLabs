using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip.ZoomLab
{
    [ExportSupertipRibbonId("ZoomLabMenu")]
    class ZoomLabMenuSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.ZoomLabMenuSupertip;
        }
    }
}
