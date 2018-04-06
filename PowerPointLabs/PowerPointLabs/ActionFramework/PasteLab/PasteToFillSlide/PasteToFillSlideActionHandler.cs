﻿using System.Windows;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportActionRibbonId(PasteLabText.PasteToFillSlideTag)]
    class PasteToFillSlideActionHandler : PasteLabActionHandler
    {
        protected override ShapeRange ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                        ShapeRange selectedShapes, ShapeRange selectedChildShapes)
        {
            ShapeRange pastingShapes = ClipboardUtil.PasteShapesFromClipboard(presentation, slide);
            if (pastingShapes == null)
            {
                Logger.Log("PasteLab: Could not paste clipboard contents.");
                MessageBox.Show(PasteLabText.ErrorPaste, PasteLabText.ErrorDialogTitle);
                return null;
            }

            PasteToFillSlide.Execute(presentation, slide, pastingShapes, presentation.SlideWidth, presentation.SlideHeight);
            return null;
        }
    }
}
