using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.EffectsLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportActionRibbonId(EffectsLabText.BlurrinessTag)]
    class EffectsLabBlurrinessActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();
            Models.PowerPointPresentation pres = this.GetCurrentPresentation();

            bool isButton = false;
            bool isCustom = ribbonId.Contains(EffectsLabText.BlurrinessCustom);
            int keywordIndex;

            if (ribbonId.Contains(CommonText.DynamicMenuButtonId))
            {
                isButton = true;
                keywordIndex = ribbonId.IndexOf(CommonText.DynamicMenuButtonId);
            }
            else
            {
                keywordIndex = ribbonId.IndexOf(CommonText.DynamicMenuOptionId);
            }

            string feature = ribbonId.Substring(0, keywordIndex);
            Selection selection = this.GetCurrentSelection();
            Models.PowerPointSlide slide = this.GetCurrentSlide();

            if (isButton)
            {
                EffectsLabSettings.ShowBlurSettingsDialog(feature);
                this.GetRibbonUi().RefreshRibbonControl(feature + CommonText.DynamicMenuOptionId + EffectsLabText.BlurrinessCustom);
            }
            else
            {
                int startIndex = keywordIndex + CommonText.DynamicMenuOptionId.Length;
                int percentage = isCustom ? GetCustomPercentage(feature) : int.Parse(ribbonId.Substring(startIndex, ribbonId.Length - startIndex));
                ExecuteBlurAction(feature, selection, pres, slide, percentage);
            }
        }

        private int GetCustomPercentage(string feature)
        {
            switch (feature)
            {
                case EffectsLabText.BlurrinessFeatureSelected:
                    return EffectsLabSettings.CustomPercentageSelected;
                case EffectsLabText.BlurrinessFeatureRemainder:
                    return EffectsLabSettings.CustomPercentageRemainder;
                case EffectsLabText.BlurrinessFeatureBackground:
                    return EffectsLabSettings.CustomPercentageBackground;
                default:
                    Logger.Log(feature + " does not exist!", Common.Logger.LogType.Error);
                    return -1;
            }
        }

        private void ExecuteBlurAction(string feature, Selection selection, Models.PowerPointPresentation pres, Models.PowerPointSlide slide, int percentage)
        {
            Utils.ClipboardUtil.RestoreClipboardAfterAction(() =>
            {
                switch (feature)
                {
                    case EffectsLabText.BlurrinessFeatureSelected:
                        EffectsLabBlur.ExecuteBlurSelected(slide, selection, percentage);
                        break;
                    case EffectsLabText.BlurrinessFeatureRemainder:
                        EffectsLabBlur.ExecuteBlurRemainder(slide, selection, percentage);
                        break;
                    case EffectsLabText.BlurrinessFeatureBackground:
                        EffectsLabBlur.ExecuteBlurBackground(slide, selection, percentage);
                        break;
                    default:
                        Logger.Log(feature + " does not exist!", Common.Logger.LogType.Error);
                        break;
                }
                return 0; // TEMPORARY
            }, pres, slide);
        }
    }
}
