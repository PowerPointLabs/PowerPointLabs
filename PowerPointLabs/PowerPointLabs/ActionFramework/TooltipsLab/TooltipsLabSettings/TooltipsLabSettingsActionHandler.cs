using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.TooltipsLab
{
    [ExportActionRibbonId(TooltipsLabText.SettingsTag)]
    class TooltipsLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            // TODO: Settings for Tooltips Lab
            MessageBox.Show("No settings for Tooltips Lab yet. Will be implemented soon!");
            return; 
        }
    }
}
