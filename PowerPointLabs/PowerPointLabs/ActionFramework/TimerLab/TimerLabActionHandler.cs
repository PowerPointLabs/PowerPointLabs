using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.TimerLab;

namespace PowerPointLabs.ActionFramework.TimerLab
{
    [ExportActionRibbonId(TimerLabText.PaneTag)]
    class TimerLabActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.RegisterTaskPane(typeof(TimerPane), TimerLabText.TaskPaneTitle);
            var timerPane = this.GetTaskPane(typeof(TimerPane));
            // if currently the pane is hidden, show the pane
            if (!timerPane.Visible)
            {
                // fire the pane visble change event
                timerPane.Visible = true;
            }
            else
            {
                timerPane.Visible = false;
            }
        }
    }
}
