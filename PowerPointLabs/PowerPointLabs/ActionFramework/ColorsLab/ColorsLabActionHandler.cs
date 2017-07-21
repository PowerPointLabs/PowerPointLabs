using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ColorsLab
{
    [ExportActionRibbonId(TextCollection1.ColorsLabTag)]
    class ColorsLabActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var colorPane = 
                this.RegisterTaskPane(typeof(ColorPane), TextCollection1.ColorsLabTaskPanelTitle);
            if (colorPane != null)
            {
                colorPane.Visible = !colorPane.Visible;
            }
        }
    }
}
