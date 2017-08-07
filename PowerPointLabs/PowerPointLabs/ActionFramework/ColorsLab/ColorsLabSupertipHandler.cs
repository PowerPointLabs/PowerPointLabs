using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ColorsLab
{
    [ExportSupertipRibbonId(ColorsLabText.PaneTag)]
    class ColorsLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return ColorsLabText.RibbonMenuSupertip;
        }
    }
}
