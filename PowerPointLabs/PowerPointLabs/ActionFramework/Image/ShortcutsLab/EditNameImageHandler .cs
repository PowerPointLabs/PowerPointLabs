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
        "EditNameMenuChart",
        "EditNameMenuTable")]
    class EditNameImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.EditNameContext);
        }
    }
}
