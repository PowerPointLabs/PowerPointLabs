using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ColorsLab
{
    [ExportActionRibbonId(ColorsLabText.PaneTag)]
    class ColorsLabActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var colorPane = 
                this.RegisterTaskPane(typeof(ColorPane), ColorsLabText.TaskPanelTitle);
            if (colorPane != null)
            {
                colorPane.Visible = !colorPane.Visible;
            }
        }
    }
}
