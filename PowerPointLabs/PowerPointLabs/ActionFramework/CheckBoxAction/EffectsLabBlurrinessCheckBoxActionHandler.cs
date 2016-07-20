using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.CheckBoxAction
{
    [ExportCheckBoxActionRibbonId(TextCollection.EffectsLabBlurrinessTag)]
    class EffectsLabBlurrinessCheckBoxActionHandler : CheckBoxActionHandler
    {
        protected override void ExecuteCheckBoxAction(string ribbonId, string ribbonTag, bool pressed)
        {
            EffectsLab.EffectsLabBlurSelected.HasOverlay = pressed;
            var ribbon = this.GetRibbonUi();
            ribbon.RefreshRibbonControl("EffectsLabBlurSelectedCheckBox");
            ribbon.RefreshRibbonControl("EffectsLabBlurRemainderCheckBox");
            ribbon.RefreshRibbonControl("EffectsLabBlurBackgroundCheckBox");
        }
    }
}
