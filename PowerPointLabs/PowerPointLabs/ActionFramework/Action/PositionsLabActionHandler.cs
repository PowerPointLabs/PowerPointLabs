using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("PositionsLabButton")]
    class PositionsLabActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.RegisterTaskPane(typeof(PositionsPane), TextCollection.PositionsLabTaskPanelTitle);
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
