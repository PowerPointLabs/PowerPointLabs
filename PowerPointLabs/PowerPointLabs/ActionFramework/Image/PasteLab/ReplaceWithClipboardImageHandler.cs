using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.PasteLab
{
    [ExportImageRibbonId(
        "ReplaceWithClipboardMenuShape", "ReplaceWithClipboardMenuLine", "ReplaceWithClipboardMenuFreeform",
        "ReplaceWithClipboardMenuPicture", "ReplaceWithClipboardMenuGroup", "ReplaceWithClipboardMenuInk",
        "ReplaceWithClipboardMenuVideo", "ReplaceWithClipboardMenuTextEdit", "ReplaceWithClipboardMenuChart",
        "ReplaceWithClipboardMenuTable", "ReplaceWithClipboardMenuTableWhole", "ReplaceWithClipboardMenuSmartArtBackground",
        "ReplaceWithClipboardButton")]
    class ReplaceWithClipboardImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.ReplaceWithClipboard);
        }
    }
}