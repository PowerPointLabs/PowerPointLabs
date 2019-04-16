using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportActionRibbonId(ShortcutsLabText.SpeakSelectedTag)]
    class SpeakSelectedActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            ComputerVoiceRuntimeService.SpeakSelectedText(
                AudioSettingService.selectedVoice as ComputerVoice);
        }
    }
}
