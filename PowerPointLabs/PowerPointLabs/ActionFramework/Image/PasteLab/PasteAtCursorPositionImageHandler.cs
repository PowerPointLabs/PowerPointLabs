using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.PasteLab
{
    [ExportImageRibbonId(
        "PasteAtCursorPosition",
        "PasteAtCursorPositionShape",
        "PasteAtCursorPositionFreeform",
        "PasteAtCursorPositionPicture",
        "PasteAtCursorPositionGroup")]
    class PasteAtCursorPositionImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.PasteLab);
        }
    }
}