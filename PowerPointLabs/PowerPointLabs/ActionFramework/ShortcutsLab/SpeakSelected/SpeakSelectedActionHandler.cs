using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.NarrationsLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportActionRibbonId(ShortcutsLabText.SpeakSelectedTag)]
    class SpeakSelectedActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            NotesToAudio.SpeakSelectedText();
        }
    }
}
