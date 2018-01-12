using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.SyncLab.Views;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.SyncLab
{
    [ExportActionRibbonId(SyncLabText.PaneTag)]
    class SyncLabActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.RegisterTaskPane(typeof(SyncPane), SyncLabText.TaskPanelTitle);
            var syncPane = this.GetTaskPane(typeof(SyncPane));
            // toggle pane visibility
            syncPane.Visible = !syncPane.Visible;
        }
    }
}
