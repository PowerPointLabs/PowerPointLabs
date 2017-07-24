using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportLabelRibbonId(PasteLabText.ReplaceWithClipboardTag)]
    class ReplaceWithClipboardLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return PasteLabText.ReplaceWithClipboardLabel;
        }
    }
}