using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.SyncLab.View;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("SyncLabButton")]
    class SyncLabActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.RegisterTaskPane(typeof(SyncPane), TextCollection.SyncLabTaskPanelTitle);
            var syncPane = this.GetTaskPane(typeof(SyncPane));
            // toggle pane visibility
            syncPane.Visible = !syncPane.Visible;
        }
    }
}
