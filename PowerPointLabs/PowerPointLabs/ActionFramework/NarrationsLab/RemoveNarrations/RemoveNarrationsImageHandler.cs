using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportImageRibbonId(TextCollection.RemoveNarrationsTag)]
    class RemoveNarrationsImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.RemoveAudio);
        }
    }
}
