﻿using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.CaptionsLab
{
    [ExportImageRibbonId(TextCollection.RemoveNotesTag)]
    class RemoveNotesImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.RemoveNotes);
        }
    }
}
