using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ResizeLab;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("ResizeLabButton")]
    class ResizeLabActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var resizePane =
                this.RegisterTaskPane(typeof(ResizeLabPane), TextCollection.ResizeLabsTaskPaneTitle);
            if (resizePane != null)
            {
                resizePane.Visible = !resizePane.Visible;
            }
        }
    }
}
