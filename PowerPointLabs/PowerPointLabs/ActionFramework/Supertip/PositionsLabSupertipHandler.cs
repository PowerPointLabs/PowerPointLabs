using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip
{
    [ExportSupertipRibbonId("PositionsLabButton")]
    class PositionsLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId, string ribbonTag)
        {
            return TextCollection.PositionsLabSupertip;
        }
    }
}
