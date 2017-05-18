using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Models;

using PPExtraEventHelper;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId("PasteAtCursorPosition")]
    class PasteAtCursorPositionActionHandler : PasteLabActionHandler
    {
        protected override void ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                    Selection selection, ShapeRange pastingShapes)
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

            if (pastingShapes.Count > 1)
            {
                Shape pastingGroup = pastingShapes.Group();
                pastingGroup.Left = positionX;
                pastingGroup.Top = positionY;
                pastingGroup.Ungroup();
            }
            else
            {
                pastingShapes[1].Left = positionX;
                pastingShapes[1].Top = positionY;
            }
        }
    }
}