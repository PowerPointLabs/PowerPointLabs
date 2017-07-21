using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Views;

namespace PowerPointLabs.ActionFramework.Help
{
    [ExportActionRibbonId(TextCollection1.AboutTag)]
    class AboutActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            AboutDialogBox dialog = new AboutDialogBox(Properties.Settings.Default.Version,
                Properties.Settings.Default.ReleaseDate, TextCollection1.PowerPointLabsWebsiteUrl);
            dialog.ShowDialog();
        }
    }
}
