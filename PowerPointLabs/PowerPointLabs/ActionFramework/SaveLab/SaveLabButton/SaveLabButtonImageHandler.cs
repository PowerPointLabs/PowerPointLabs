using System.Drawing;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.SaveLab
{
    [ExportImageRibbonId(SaveLabText.SavePresentationsButtonTag)]
    class SaveLabButtonImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            // Need a new image for SaveLab
            return new Bitmap(Properties.Resources.SaveLabActionButton);
        }
    }
}
