using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.PasteLab
{
    [ExportImageRibbonId(
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuShape,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuLine,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuFreeform,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuPicture,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuGroup,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuInk,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuVideo,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuChart,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuTable,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuTableCell,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.MenuSmartArt,
        TextCollection.ReplaceWithClipboardMenuId + TextCollection.RibbonButton)]
    class ReplaceWithClipboardImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.ReplaceWithClipboard);
        }
    }
}