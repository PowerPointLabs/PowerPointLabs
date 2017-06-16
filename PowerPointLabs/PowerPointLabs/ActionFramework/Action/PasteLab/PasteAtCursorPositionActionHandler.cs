using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Models;
using PowerPointLabs.PasteLab;

using PPExtraEventHelper;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId(
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuShape,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuLine,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuFreeform,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuPicture,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuGroup,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuInk,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuVideo,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuTextEdit,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuChart,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuTable,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuTableCell,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuSlide,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuSmartArt,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuEditSmartArt,
        TextCollection.PasteAtCursorPositionMenuId + TextCollection.MenuEditSmartArtText)]
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