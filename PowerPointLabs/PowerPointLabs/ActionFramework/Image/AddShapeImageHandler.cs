using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId("AddCustomShape", "AddCustomShapePicture", "AddCustomShapeChart", 
                        "AddCustomShapeTable", "AddCustomShapeGroup", "AddCustomShapeFreeform",
                        "AddCustomShapeInk", "AddCustomShapeSmartArt")]
    class AddShapeImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.AddToCustomShapes);
        }
    }
}
