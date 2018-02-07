using System.Drawing;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportImageRibbonId(SaveLabText.RibbonMenuId)]
    class SaveLabMenuImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            // Need a new icon for Save Lab
            return new Bitmap(Properties.Resources.CropLab);
        }
    }
}
