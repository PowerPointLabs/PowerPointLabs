using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.PasteLab
{
    [ExportImageRibbonId(
        "ReplaceWithClipboardMenuShape",
        "ReplaceWithClipboardMenuLine",
        "ReplaceWithClipboardMenuFreeform",
        "ReplaceWithClipboardMenuPicture",
        "ReplaceWithClipboardMenuGroup",
        "ReplaceWithClipboardMenuChart",
        "ReplaceWithClipboardMenuTable",
        "ReplaceWithClipboardMenuTableWhole")]
    class ReplaceWithClipboardImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.PasteLab);
        }
    }
}