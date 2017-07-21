using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.EffectsLab;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportActionRibbonId(TextCollection1.EffectsLabBlurrinessTag)]
    class EffectsLabBlurrinessActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            bool isButton = false;
            bool isCustom = ribbonId.Contains(TextCollection1.EffectsLabBlurrinessCustom);
            int keywordIndex;

            if (ribbonId.Contains(TextCollection1.DynamicMenuButtonId))
            {
                isButton = true;
                keywordIndex = ribbonId.IndexOf(TextCollection1.DynamicMenuButtonId);
            }
            else
            {
                keywordIndex = ribbonId.IndexOf(TextCollection1.DynamicMenuOptionId);
            }

            string feature = ribbonId.Substring(0, keywordIndex);
            Selection selection = this.GetCurrentSelection();
            Models.PowerPointSlide slide = this.GetCurrentSlide();

            if (isButton)
            {
                EffectsLabSettings.ShowBlurSettingsDialog(feature);
                this.GetRibbonUi().RefreshRibbonControl(feature + TextCollection1.DynamicMenuOptionId + TextCollection1.EffectsLabBlurrinessCustom);
            }
            else
            {
                int startIndex = keywordIndex + TextCollection1.DynamicMenuOptionId.Length;
                int percentage = isCustom ? GetCustomPercentage(feature) : int.Parse(ribbonId.Substring(startIndex, ribbonId.Length - startIndex));
                ExecuteBlurAction(feature, selection, slide, percentage);
            }
        }

        private int GetCustomPercentage(string feature)
        {
            switch (feature)
            {
                case TextCollection1.EffectsLabBlurrinessFeatureSelected:
                    return EffectsLabSettings.CustomPercentageSelected;
                case TextCollection1.EffectsLabBlurrinessFeatureRemainder:
                    return EffectsLabSettings.CustomPercentageRemainder;
                case TextCollection1.EffectsLabBlurrinessFeatureBackground:
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
                case TextCollection1.EffectsLabBlurrinessFeatureSelected:
                    EffectsLabBlur.ExecuteBlurSelected(slide, selection, percentage);
                    break;
                case TextCollection1.EffectsLabBlurrinessFeatureRemainder:
                    EffectsLabBlur.ExecuteBlurRemainder(slide, selection, percentage);
                    break;
                case TextCollection1.EffectsLabBlurrinessFeatureBackground:
                    EffectsLabBlur.ExecuteBlurBackground(slide, selection, percentage);
                    break;
                default:
                    Logger.Log(feature + " does not exist!", Common.Logger.LogType.Error);
                    break;
            }
        }
    }
}
