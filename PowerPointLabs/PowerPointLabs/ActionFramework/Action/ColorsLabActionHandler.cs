using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("ColorsLabButton")]
    class ColorsLabActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var colorPane = 
                this.RegisterTaskPane(typeof(ColorPane), TextCollection.ColorsLabTaskPanelTitle);
            if (colorPane != null)
            {
                colorPane.Visible = !colorPane.Visible;
            }
        }
    }
}
