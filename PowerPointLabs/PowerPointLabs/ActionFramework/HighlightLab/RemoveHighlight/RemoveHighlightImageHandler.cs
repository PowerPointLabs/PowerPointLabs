using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportImageRibbonId(TextCollection.RemoveHighlightTag)]
    class RemoveHighlightImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.RemoveHighlighting);
        }
    }
}
