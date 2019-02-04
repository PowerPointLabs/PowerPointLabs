using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.NarrationsLab;
using PowerPointLabs.NarrationsLab.Storage;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportActionRibbonId(NarrationsLabText.SettingsTag)]
    class NarrationsLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            NarrationsLabStorageConfig.LoadUserAccount();
            NarrationsLabSettings.ShowSettingsDialog();
        }
    }
}
