using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId(
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuShape,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuLine,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuFreeform,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuPicture,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuGroup,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuInk,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuVideo,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuTextEdit,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuChart,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuTable,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuTableCell,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuSmartArt,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuEditSmartArt,
        TextCollection.AddCustomShapeMenuId + TextCollection.MenuEditSmartArtText)]
    class AddShapeImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.AddToCustomShapes);
        }
    }
}
