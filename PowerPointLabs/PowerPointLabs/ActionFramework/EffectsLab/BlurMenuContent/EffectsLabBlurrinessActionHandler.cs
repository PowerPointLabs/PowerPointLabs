using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.EffectsLab;
using PowerPointLabs.EffectsLab.Views;

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
                EffectsLabBlurDialogBox dialog = new EffectsLabBlurDialogBox(feature);
                dialog.SettingsHandler += PropertiesEdited;
                dialog.ShowDialog();
            }
            else
            {
                int startIndex = keywordIndex + TextCollection.DynamicMenuOptionId.Length;
                int percentage = isCustom ? GetCustomPercentage(feature) : int.Parse(ribbonId.Substring(startIndex, ribbonId.Length - startIndex));
                ExecuteBlurAction(feature, selection, slide, percentage);
            }
        }

        private void PropertiesEdited(string feature, int percentage, bool isTint)
        {
            switch (feature)
            {
                case TextCollection.EffectsLabBlurrinessFeatureSelected:
                    EffectsLabBlur.IsTintSelected = isTint;
                    EffectsLabBlur.CustomPercentageSelected = percentage;
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureRemainder:
                    EffectsLabBlur.IsTintRemainder = isTint;
                    EffectsLabBlur.CustomPercentageRemainder = percentage;
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureBackground:
                    EffectsLabBlur.IsTintBackground = isTint;
                    EffectsLabBlur.CustomPercentageBackground = percentage;
                    break;
                default:
                    Logger.Log(feature + " does not exist!", Common.Logger.LogType.Error);
                    break;
            }
            
            this.GetRibbonUi().RefreshRibbonControl(feature + TextCollection.DynamicMenuOptionId + TextCollection.EffectsLabBlurrinessCustom);
        }

        private int GetCustomPercentage(string feature)
        {
            switch (feature)
            {
                case TextCollection.EffectsLabBlurrinessFeatureSelected:
                    return EffectsLabBlur.CustomPercentageSelected;
                case TextCollection.EffectsLabBlurrinessFeatureRemainder:
                    return EffectsLabBlur.CustomPercentageRemainder;
                case TextCollection.EffectsLabBlurrinessFeatureBackground:
                    return EffectsLabBlur.CustomPercentageBackground;
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
