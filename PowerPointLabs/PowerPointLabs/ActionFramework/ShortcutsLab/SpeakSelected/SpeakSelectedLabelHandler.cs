using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportLabelRibbonId(TextCollection.SpeakSelectedTag)]
    class SpeakSelectedLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.ContextSpeakSelectedTextLabel;
        }
    }
}
