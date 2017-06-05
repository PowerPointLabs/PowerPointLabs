using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("ShapesLabButton")]
    class ShapesMenuActionHandler : ShapesLabActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            InitCustomShapePane();
            TogglePaneVisibility();
        }
    }
}
