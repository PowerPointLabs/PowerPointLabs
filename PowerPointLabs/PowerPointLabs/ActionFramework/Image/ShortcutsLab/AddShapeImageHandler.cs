using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId(
        TextCollection.AddCustomShapeId + TextCollection.MenuShape,
        TextCollection.AddCustomShapeId + TextCollection.MenuLine,
        TextCollection.AddCustomShapeId + TextCollection.MenuFreeform,
        TextCollection.AddCustomShapeId + TextCollection.MenuPicture,
        TextCollection.AddCustomShapeId + TextCollection.MenuGroup,
        TextCollection.AddCustomShapeId + TextCollection.MenuInk,
        TextCollection.AddCustomShapeId + TextCollection.MenuVideo,
        TextCollection.AddCustomShapeId + TextCollection.MenuTextEdit,
        TextCollection.AddCustomShapeId + TextCollection.MenuChart,
        TextCollection.AddCustomShapeId + TextCollection.MenuTable,
        TextCollection.AddCustomShapeId + TextCollection.MenuTableCell,
        TextCollection.AddCustomShapeId + TextCollection.MenuSmartArt,
        TextCollection.AddCustomShapeId + TextCollection.MenuEditSmartArt,
        TextCollection.AddCustomShapeId + TextCollection.MenuEditSmartArtText)]
    class AddShapeImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.AddToCustomShapes);
        }
    }
}
