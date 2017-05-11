using System.Windows;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PPExtraEventHelper;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId("pasteToPosition")]
    class PasteLabPositionActionHandler : PasteLabActionHandler
    {
        protected override void ExecutePasteAction(string ribbonId)
        {
            var slide = this.GetCurrentSlide();
            var coordinates = PPMouse.RightClickCoordinates;
            var activeWindow = this.GetCurrentWindow();
            bool clipboardIsEmpty = (Clipboard.GetDataObject() == null);

            float xPosition = 0;
            float yPosition = 0;

            if (activeWindow.ActivePane.ViewType == PowerPoint.PpViewType.ppViewSlide)
            {
                int xref = activeWindow.PointsToScreenPixelsX(100) - activeWindow.PointsToScreenPixelsX(0);
                int yref = activeWindow.PointsToScreenPixelsY(100) - activeWindow.PointsToScreenPixelsY(0);
                xPosition = ((float)(coordinates.X - activeWindow.PointsToScreenPixelsX(0)) / (float)xref) * 100;
                yPosition = ((float)(coordinates.Y - activeWindow.PointsToScreenPixelsY(0)) / (float)yref) * 100;
            }

            PowerPointLabs.PasteLab.PasteLabMain.PasteToPosition(slide, clipboardIsEmpty, xPosition, yPosition);
        }
    }
}