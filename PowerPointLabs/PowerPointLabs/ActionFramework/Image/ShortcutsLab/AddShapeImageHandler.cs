using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId("AddToShapesLabMenuShape", "AddToShapesLabMenuLine", "AddToShapesLabMenuPicture",
                        "AddToShapesLabMenuChart", "AddToShapesLabMenuTable", "AddToShapesLabMenuGroup", "AddToShapesLabMenuVideo",
                        "AddToShapesLabMenuTextEdit", "AddToShapesLabMenuFreeform", "AddToShapesLabMenuInk", "AddToShapesLabMenuSmartArtBackground",
                        "AddToShapesLabMenuTableWhole", "AddToShapesLabMenuSmartArtEditSmartArt", "AddToShapesLabMenuSmartArtEditText")]
    class AddShapeImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.AddToCustomShapes);
        }
    }
}
