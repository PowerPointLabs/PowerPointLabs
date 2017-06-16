using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.PasteLab
{
    [ExportImageRibbonId(
        TextCollection.PasteIntoGroupId + TextCollection.MenuGroup,
        TextCollection.PasteIntoGroupId + TextCollection.RibbonButton)]
    class PasteIntoGroupImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.PasteIntoGroup);
        }
    }
}