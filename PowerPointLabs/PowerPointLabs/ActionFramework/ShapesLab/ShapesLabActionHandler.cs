using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ShapesLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShapesLab
{
    [ExportActionRibbonId(ShapesLabText.PaneTag)]
    class ShapesLabActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.RegisterTaskPane(typeof(CustomShapePane), ShapesLabText.TaskPanelTitle);
            CustomTaskPane customShapePane = this.GetTaskPane(typeof(CustomShapePane));
            // toggle pane visibility
            customShapePane.Visible = !customShapePane.Visible;
        }
    }
}
