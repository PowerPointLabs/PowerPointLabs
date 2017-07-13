using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.NarrationsLab;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportActionRibbonId(TextCollection.SpeakSelectedTag)]
    class SpeakSelectedActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            NotesToAudio.SpeakSelectedText();
        }
    }
}
