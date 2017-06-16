using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;

using PPExtraEventHelper;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId(
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuShape,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuLine,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuFreeform,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuPicture,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuGroup,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuInk,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuVideo,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuTextEdit,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuChart,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuTable,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuTableCell,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuSlide,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuSmartArt,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuEditSmartArt,
        TextCollection.PasteAtCursorPositionId + TextCollection.MenuEditSmartArtText)]
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

            ShapeRange pastingShapes = PasteShapesFromClipboard(slide);
            if (pastingShapes == null)
            {
                return null;
            }

            return PasteAtCursorPosition.Execute(presentation, slide, pastingShapes, positionX, positionY);
        }
    }
}