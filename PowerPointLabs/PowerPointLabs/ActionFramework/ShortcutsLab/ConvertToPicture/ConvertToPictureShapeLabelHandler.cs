﻿using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportLabelRibbonId(TextCollection.ConvertToPictureTag)]
    class ConvertToPictureShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.ConvertToPictureShapeLabel;
        }
    }
}
