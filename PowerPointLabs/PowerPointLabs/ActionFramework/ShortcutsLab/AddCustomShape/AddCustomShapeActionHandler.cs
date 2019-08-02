using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;
using PowerPointLabs.ShapesLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportActionRibbonId(ShortcutsLabText.AddCustomShapeTag)]
    class AddCustomShapeActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();
            PowerPointPresentation pres = this.GetCurrentPresentation();
            PowerPointSlide slide = this.GetCurrentSlide();
            Selection selection = this.GetCurrentSelection();

            CustomTaskPane shapesLabPane = this.GetTaskPane(typeof(CustomShapePane));

            if (shapesLabPane == null)
            {
                this.RegisterTaskPane(typeof(CustomShapePane), ShapesLabText.TaskPanelTitle);
                shapesLabPane = this.GetTaskPane(typeof(CustomShapePane));
            }
            if (shapesLabPane == null)
            {
                return;
            }
            shapesLabPane.Visible = true;

            CustomShapePane customShapePane = shapesLabPane.Control as CustomShapePane;
            customShapePane.InitCustomShapePaneStorage();
            customShapePane.AddShapeFromSelection(selection);
        }
    }
}
