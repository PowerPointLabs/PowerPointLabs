using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.NarrationsLab;

namespace PowerPointLabs.ActionFramework.Animationlab
{
    [ExportActionRibbonId(TextCollection.NarrationsLabSettingsTag)]
    class NarrationsLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            NarrationsLabSettings.ShowSettingsDialog();
        }
    }
}
