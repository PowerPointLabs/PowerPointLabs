﻿using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.HighlightLab
{
    [ExportImageRibbonId(TextCollection.HighlightLabMenuId)]
    class HighlightLabMenuImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.HighlightLab);
        }
    }
}
