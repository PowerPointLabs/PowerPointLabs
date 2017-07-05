using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.RibbonGroups
{
    [ExportImageRibbonId(TextCollection.FormattingGroupId)]
    class FormattingGroupImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.FormattingGroup);
        }
    }
}
