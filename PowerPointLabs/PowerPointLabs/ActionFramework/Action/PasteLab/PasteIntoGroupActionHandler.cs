﻿using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId("PasteIntoGroup")]
    class PasteIntoGroupActionHandler : PasteLabActionHandler
    {
        protected override void ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                    Selection selection, ShapeRange pastingShapes)
        {
            if (!IsSelectionShapes(selection))
            {
                Logger.Log("PasteIntoGroup failed. No valid shape is selected.");
                return;
            }

            if (selection.ShapeRange.Count == 1 && !Graphics.IsAGroup(selection.ShapeRange[1]))
            {
                Logger.Log("PasteIntoGroup failed. Selection is only a single shape.");
                return;
            }

            PasteIntoGroup.Execute(presentation, slide, selection.ShapeRange, pastingShapes);
        }
    }
}