using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.ELearningLab.Service.StorageService;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Views;

namespace PowerPointLabs.ActionFramework.ELearningLab
{
    [ExportActionRibbonId(ELearningLabText.ELearningLabSettingsTag)]
    class ELearningLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            LoadingDialogBox splashView = new LoadingDialogBox();
            splashView.Show();
            AzureAccountStorageService.LoadUserAccount();
            WatsonAccountStorageService.LoadUserAccount();
            AudioSettingStorageService.LoadAudioSettingPreference();
            splashView.Close();
            AudioSettingService.ShowSettingsDialog();
        }
    }
}
