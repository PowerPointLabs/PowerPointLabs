using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportEnabledRibbonId(NarrationsLabText.RemoveNarrationsTag)]
    class RemoveNarrationsEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return ComputerVoiceRuntimeService.IsRemoveAudioEnabled;
        }
    }
}