using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId(
        "EditNameMenuShape",
        "EditNameMenuLine",
        "EditNameMenuFreeform",
        "EditNameMenuPicture",
        "EditNameMenuGroup",
        "EditNameMenuInk",
        "EditNameMenuVideo",
        "EditNameMenuTextEdit",
        "EditNameMenuChart",
        "EditNameMenuTable",
        "EditNameMenuTableWhole",
        "EditNameMenuSmartArtBackground",
        "EditNameMenuSmartArtEditSmartArt",
        "EditNameMenuSmartArtEditText")]
    class EditNameImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.EditNameContext);
        }
    }
}
