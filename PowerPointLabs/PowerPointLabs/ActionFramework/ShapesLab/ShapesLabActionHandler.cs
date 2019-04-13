using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
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
            CustomTaskPane shapesLabPane = this.GetTaskPane(typeof(CustomShapePane));

            if (shapesLabPane == null)
            {
                return;
            }
            
            // toggle pane visibility
            shapesLabPane.Visible = !shapesLabPane.Visible;

            CustomShapePane customShapePane = shapesLabPane.Control as CustomShapePane;
            customShapePane.InitCustomShapePaneStorage();
        }
    }
}
