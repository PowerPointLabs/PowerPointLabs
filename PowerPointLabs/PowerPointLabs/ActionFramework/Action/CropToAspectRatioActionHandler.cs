using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.CropLab;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId(TextCollection.CropToAspectRatioTag)]
    class CropToAspectRatioActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            if (ribbonId.Contains(TextCollection.DynamicMenuButtonId))
            {
                var dialog = new CropLab.CustomAspectRatioDialog();
                dialog.SettingsHandler += ExecuteCropToAspectRatio;
                dialog.ShowDialog();
            }
            else if (ribbonId.Contains(TextCollection.DynamicMenuOptionId))
            {
                int optionRawStringStartIndex = ribbonId.LastIndexOf(TextCollection.DynamicMenuButtonId) +
                                                TextCollection.DynamicMenuOptionId.Length;
                string optionRawString = ribbonId.Substring(optionRawStringStartIndex).Replace('_', ':');
                ExecuteCropToAspectRatio(optionRawString);
            }
        }

        private void ExecuteCropToAspectRatio(string aspectRatioRawString)
        {
            this.StartNewUndoEntry();
            var selection = this.GetCurrentSelection();
            CropLabErrorHandler errorHandler = CropLabErrorHandler.InitializeErrorHandler(CropLabUIControl.GetSharedInstance());
            CropToAspectRatio.Crop(selection, aspectRatioRawString, errorHandler);
        }
    }
}
