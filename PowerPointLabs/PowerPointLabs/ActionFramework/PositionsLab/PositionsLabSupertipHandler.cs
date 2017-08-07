using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PositionsLab
{
    [ExportSupertipRibbonId(PositionsLabText.PaneTag)]
    class PositionsLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return PositionsLabText.RibbonMenuSupertip;
        }
    }
}
