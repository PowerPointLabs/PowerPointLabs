using System.Drawing;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.HighlightLab;
using PowerPointLabs.Views;

namespace PowerPointLabs.ActionFramework.Highlightlab
{
    [ExportActionRibbonId(TextCollection.HighlightLabSettingsTag)]
    class HighlightLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var dialog = new HighlightLabSettingsDialogBox(HighlightBulletsText.highlightColor,
                HighlightBulletsText.defaultColor, HighlightBulletsBackground.backgroundColor);
            dialog.SettingsHandler += HighlightBulletsPropertiesEdited;
            dialog.ShowDialog();
        }

        private void HighlightBulletsPropertiesEdited(Color newHighlightColor, Color newDefaultColor, Color newBackgroundColor)
        {
            HighlightBulletsText.highlightColor = newHighlightColor;
            HighlightBulletsText.defaultColor = newDefaultColor;
            HighlightBulletsBackground.backgroundColor = newBackgroundColor;
            HighlightTextFragments.backgroundColor = newBackgroundColor;
        }
    }
}
