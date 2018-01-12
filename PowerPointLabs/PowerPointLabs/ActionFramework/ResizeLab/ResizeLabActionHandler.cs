using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ResizeLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ResizeLab
{
    [ExportActionRibbonId(ResizeLabText.PaneTag)]
    class ResizeLabActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var resizePane =
                this.RegisterTaskPane(typeof(ResizeLabPane), ResizeLabText.TaskPaneTitle);
            if (resizePane != null)
            {
                resizePane.Visible = !resizePane.Visible;
            }
        }
    }
}
