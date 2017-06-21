using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportImageRibbonId(TextCollection.ReplaceWithClipboardTag)]
    class ReplaceWithClipboardImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.ReplaceWithClipboard);
        }
    }
}