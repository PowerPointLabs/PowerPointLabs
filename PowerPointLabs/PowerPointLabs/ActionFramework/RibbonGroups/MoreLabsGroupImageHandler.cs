using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.RibbonGroups
{
    [ExportImageRibbonId(TextCollection.MoreLabsGroupId)]
    class MoreLabsGroupImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.PptlabsContextMenu);
        }
    }
}
