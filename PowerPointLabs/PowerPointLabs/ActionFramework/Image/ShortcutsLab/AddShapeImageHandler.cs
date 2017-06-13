using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId(
        "AddToShapesLabMenuShape", "AddToShapesLabMenuLine", "AddToShapesLabMenuFreeform",
        "AddToShapesLabMenuPicture", "AddToShapesLabMenuGroup", "AddToShapesLabMenuInk",
        "AddToShapesLabMenuVideo", "AddToShapesLabMenuTextEdit", "AddToShapesLabMenuChart",
        "AddToShapesLabMenuTable", "AddToShapesLabMenuTableWhole", "AddToShapesLabMenuSmartArtBackground",
        "AddToShapesLabMenuSmartArtEditSmartArt", "AddToShapesLabMenuSmartArtEditText")]
    class AddShapeImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.AddToCustomShapes);
        }
    }
}
