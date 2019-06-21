using System.Windows;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;
using PPExtraEventHelper;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportActionRibbonId(PasteLabText.PasteAtCursorPositionTag)]
    class PasteAtCursorPositionActionHandler : PasteLabActionHandler
    {
        protected override ShapeRange ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                        ShapeRange selectedShapes, ShapeRange selectedChildShapes)
        {
            PPMouse.Coordinates coordinates = PPMouse.RightClickCoordinates;
            DocumentWindow activeWindow = this.GetCurrentWindow();
            
            float positionX = 0;
            float positionY = 0;

            if (activeWindow.ActivePane.ViewType == PpViewType.ppViewSlide)
            {
                int xref = activeWindow.PointsToScreenPixelsX(100) - activeWindow.PointsToScreenPixelsX(0);
                int yref = activeWindow.PointsToScreenPixelsY(100) - activeWindow.PointsToScreenPixelsY(0);
                positionX = ((coordinates.X - activeWindow.PointsToScreenPixelsX(0)) / xref) * 100;
                positionY = ((coordinates.Y - activeWindow.PointsToScreenPixelsY(0)) / yref) * 100;
            }

            ShapeRange pastingShapes = ClipboardUtil.PasteShapesFromClipboard(presentation, slide);
            if (pastingShapes == null)
            {
                Logger.Log("PasteLab: Could not paste clipboard contents.");
                MessageBoxUtil.Show(PasteLabText.ErrorPaste, PasteLabText.ErrorDialogTitle);
                return null;
            }

            return PasteAtCursorPosition.Execute(presentation, slide, pastingShapes, positionX, positionY);
        }
    }
}