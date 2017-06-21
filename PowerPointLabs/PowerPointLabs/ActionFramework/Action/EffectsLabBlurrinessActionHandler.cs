using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.EffectsLab;
using PowerPointLabs.EffectsLab.Views;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId(TextCollection.EffectsLabBlurrinessTag)]
    class EffectsLabBlurrinessActionHandler : ActionHandler
    {
        private string feature;
        private Microsoft.Office.Interop.PowerPoint.Selection selection;
        private Models.PowerPointSlide slide;

        protected override void ExecuteAction(string ribbonId)
        {
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

            feature = ribbonId.Substring(0, keywordIndex);
            selection = this.GetCurrentSelection();
            slide = this.GetCurrentSlide();

            if (isButton)
            {
                EffectsLabBlurDialogBox dialog = new EffectsLabBlurDialogBox(feature);
                dialog.SettingsHandler += PropertiesEdited;
                dialog.ShowDialog();
            }
            else
            {
                int startIndex = keywordIndex + TextCollection.DynamicMenuOptionId.Length;
                int percentage = isCustom ? GetCustomPercentage() : int.Parse(ribbonId.Substring(startIndex, ribbonId.Length - startIndex));
                ExecuteBlurAction(percentage);
            }
        }

        private void PropertiesEdited(int percentage, bool isTint)
        {
            switch (feature)
            {
                case TextCollection.EffectsLabBlurrinessFeatureSelected:
                    EffectsLabBlurSelected.IsTintSelected = isTint;
                    EffectsLabBlurSelected.CustomPercentageSelected = percentage;
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureRemainder:
                    EffectsLabBlurSelected.IsTintRemainder = isTint;
                    EffectsLabBlurSelected.CustomPercentageRemainder = percentage;
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureBackground:
                    EffectsLabBlurSelected.IsTintBackground = isTint;
                    EffectsLabBlurSelected.CustomPercentageBackground = percentage;
                    break;
                default:
                    throw new System.Exception("Invalid feature");
            }
            
            this.GetRibbonUi().RefreshRibbonControl(feature + TextCollection.DynamicMenuOptionId + TextCollection.EffectsLabBlurrinessCustom);
        }

        private int GetCustomPercentage()
        {
            switch (feature)
            {
                case TextCollection.EffectsLabBlurrinessFeatureSelected:
                    return EffectsLabBlurSelected.CustomPercentageSelected;
                case TextCollection.EffectsLabBlurrinessFeatureRemainder:
                    return EffectsLabBlurSelected.CustomPercentageRemainder;
                case TextCollection.EffectsLabBlurrinessFeatureBackground:
                    return EffectsLabBlurSelected.CustomPercentageBackground;
                default:
                    throw new System.Exception("Invalid feature");
            }
        }

        private void ExecuteBlurAction(int percentage)
        {
            switch (feature)
            {
                case TextCollection.EffectsLabBlurrinessFeatureSelected:
                    this.StartNewUndoEntry();
                    EffectsLabBlurSelected.BlurSelected(slide, selection, percentage);
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureRemainder:
                    this.GetRibbonUi().BlurRemainderEffectClick(percentage);
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureBackground:
                    this.GetRibbonUi().BlurBackgroundEffectClick(percentage);
                    break;
                default:
                    throw new System.Exception("Invalid feature");
            }
        }
    }
}
