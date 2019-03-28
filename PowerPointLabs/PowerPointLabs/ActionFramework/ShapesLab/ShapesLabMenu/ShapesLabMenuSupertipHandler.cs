using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShapesLab.ShapesLabMenu
{
    [ExportSupertipRibbonId(ShapesLabText.RibbonMenuId)]
    class ShapesLabMenuSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return ShapesLabText.RibbonMenuSupertip;
        }
    }
}