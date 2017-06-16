using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId(
        TextCollection.EditNameId + TextCollection.MenuShape,
        TextCollection.EditNameId + TextCollection.MenuLine,
        TextCollection.EditNameId + TextCollection.MenuFreeform,
        TextCollection.EditNameId + TextCollection.MenuPicture,
        TextCollection.EditNameId + TextCollection.MenuGroup,
        TextCollection.EditNameId + TextCollection.MenuInk,
        TextCollection.EditNameId + TextCollection.MenuVideo,
        TextCollection.EditNameId + TextCollection.MenuTextEdit,
        TextCollection.EditNameId + TextCollection.MenuChart,
        TextCollection.EditNameId + TextCollection.MenuTable,
        TextCollection.EditNameId + TextCollection.MenuTableCell,
        TextCollection.EditNameId + TextCollection.MenuSmartArt,
        TextCollection.EditNameId + TextCollection.MenuEditSmartArt,
        TextCollection.EditNameId + TextCollection.MenuEditSmartArtText)]
    class EditNameImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.EditNameContext);
        }
    }
}
