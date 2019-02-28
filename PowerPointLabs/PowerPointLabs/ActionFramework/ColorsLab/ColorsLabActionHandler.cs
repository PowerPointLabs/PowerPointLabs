using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ColorsLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ColorsLab
{
    [ExportActionRibbonId(ColorsLabText.PaneTag)]
    class ColorsLabActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.RegisterTaskPane(typeof(ColorsLabPane), ColorsLabText.TaskPanelTitle);
            CustomTaskPane colorPane = this.GetTaskPane(typeof(ColorsLabPane));

            // if currently the pane is hidden, show the pane, vice versa.
            colorPane.Visible = !colorPane.Visible;
        }
    }
}
