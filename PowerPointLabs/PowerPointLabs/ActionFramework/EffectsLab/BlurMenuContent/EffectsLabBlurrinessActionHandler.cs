using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.EffectsLab;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportActionRibbonId(TextCollection.EffectsLabBlurrinessTag)]
    class EffectsLabBlurrinessActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            bool isButton = false;
            bool isCustom = ribbonId.Contains(TextCollection.EffectsLabBlurrinessCustom);
            int keywordIndex;

            if (ribbonId.Contains(TextCollection.DynamicMenuButtonId))
            {
                isButton = true;
                keywordIndex = ribbonId.IndexOf(TextCollection.DynamicMenuButtonId);
            }
            else
            {
                keywordIndex = ribbonId.IndexOf(TextCollection.DynamicMenuOptionId);
            }

            string feature = ribbonId.Substring(0, keywordIndex);
            Selection selection = this.GetCurrentSelection();
            Models.PowerPointSlide slide = this.GetCurrentSlide();

            if (isButton)
            {
                EffectsLabSettings.ShowBlurSettingsDialog(feature);
                this.GetRibbonUi().RefreshRibbonControl(feature + TextCollection.DynamicMenuOptionId + TextCollection.EffectsLabBlurrinessCustom);
            }
            else
            {
                int startIndex = keywordIndex + TextCollection.DynamicMenuOptionId.Length;
                int percentage = isCustom ? GetCustomPercentage(feature) : int.Parse(ribbonId.Substring(startIndex, ribbonId.Length - startIndex));
                ExecuteBlurAction(feature, selection, slide, percentage);
            }
        }

        private int GetCustomPercentage(string feature)
        {
            switch (feature)
            {
                case TextCollection.EffectsLabBlurrinessFeatureSelected:
                    return EffectsLabSettings.CustomPercentageSelected;
                case TextCollection.EffectsLabBlurrinessFeatureRemainder:
                    return EffectsLabSettings.CustomPercentageRemainder;
                case TextCollection.EffectsLabBlurrinessFeatureBackground:
                    return EffectsLabSettings.CustomPercentageBackground;
                default:
                    Logger.Log(feature + " does not exist!", Common.Logger.LogType.Error);
                    return -1;
            }
        }

        private void ExecuteBlurAction(string feature, Selection selection, Models.PowerPointSlide slide, int percentage)
        {
            switch (feature)
            {
                case TextCollection.EffectsLabBlurrinessFeatureSelected:
                    EffectsLabBlur.ExecuteBlurSelected(slide, selection, percentage);
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureRemainder:
                    EffectsLabBlur.ExecuteBlurRemainder(slide, selection, percentage);
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureBackground:
                    EffectsLabBlur.ExecuteBlurBackground(slide, selection, percentage);
                    break;
                default:
                    Logger.Log(feature + " does not exist!", Common.Logger.LogType.Error);
                    break;
            }
        }
    }
}
