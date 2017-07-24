using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PositionsLab
{
    [ExportActionRibbonId(PositionsLabText.PaneTag)]
    class PositionsLabActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.RegisterTaskPane(typeof(PositionsPane), PositionsLabText.TaskPanelTitle);
            var positionsPane = this.GetTaskPane(typeof(PositionsPane));
            // if currently the pane is hidden, show the pane
            if (!positionsPane.Visible)
            {
                // fire the pane visble change event
                positionsPane.Visible = true;
            }
            else
            {
                positionsPane.Visible = false;
            }
        }
    }
}
