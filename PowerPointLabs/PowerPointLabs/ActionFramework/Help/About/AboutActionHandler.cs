using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Views;

namespace PowerPointLabs.ActionFramework.Help
{
    [ExportActionRibbonId(HelpText.AboutTag)]
    class AboutActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            AboutDialogBox dialog = new AboutDialogBox(Properties.Settings.Default.Version,
                Properties.Settings.Default.ReleaseDate, CommonText.PowerPointLabsWebsiteUrl);
            dialog.ShowThematicDialog();
        }
    }
}
