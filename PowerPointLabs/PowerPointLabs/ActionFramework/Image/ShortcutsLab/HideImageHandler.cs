using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId(
        TextCollection.HideSelectedShapeId + TextCollection.MenuShape,
        TextCollection.HideSelectedShapeId + TextCollection.MenuLine,
        TextCollection.HideSelectedShapeId + TextCollection.MenuFreeform,
        TextCollection.HideSelectedShapeId + TextCollection.MenuPicture,
        TextCollection.HideSelectedShapeId + TextCollection.MenuGroup,
        TextCollection.HideSelectedShapeId + TextCollection.MenuInk,
        TextCollection.HideSelectedShapeId + TextCollection.MenuVideo,
        TextCollection.HideSelectedShapeId + TextCollection.MenuTextEdit,
        TextCollection.HideSelectedShapeId + TextCollection.MenuChart,
        TextCollection.HideSelectedShapeId + TextCollection.MenuTable,
        TextCollection.HideSelectedShapeId + TextCollection.MenuTableCell,
        TextCollection.HideSelectedShapeId + TextCollection.MenuSmartArt,
        TextCollection.HideSelectedShapeId + TextCollection.MenuEditSmartArt,
        TextCollection.HideSelectedShapeId + TextCollection.MenuEditSmartArtText)]
    class HideImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.HideShape);
        }
    }
}
