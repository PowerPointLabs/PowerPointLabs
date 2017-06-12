using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.PasteLab
{
    [ExportLabelRibbonId(
        "ReplaceWithClipboardMenuShape",
        "ReplaceWithClipboardMenuLine",
        "ReplaceWithClipboardMenuFreeform",
        "ReplaceWithClipboardMenuPicture",
        "ReplaceWithClipboardMenuGroup",
        "ReplaceWithClipboardMenuChart",
        "ReplaceWithClipboardMenuTable",
        "ReplaceWithClipboardMenuTableWhole")]
    class ReplaceWithClipboardLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.PasteLabText.ReplaceWithClipboard;
        }
    }
}