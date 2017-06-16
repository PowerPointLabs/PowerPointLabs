using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.PasteLab
{
    [ExportImageRibbonId(
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuShape,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuLine,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuFreeform,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuPicture,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuGroup,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuInk,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuVideo,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuChart,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuTable,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuTableCell,
        TextCollection.ReplaceWithClipboardId + TextCollection.MenuSmartArt,
        TextCollection.ReplaceWithClipboardId + TextCollection.RibbonButton)]
    class ReplaceWithClipboardImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.ReplaceWithClipboard);
        }
    }
}