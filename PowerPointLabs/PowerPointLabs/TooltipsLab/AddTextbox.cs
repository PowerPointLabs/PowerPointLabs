using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.TooltipsLab;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.TooltipsLab
{
    internal static class AddTextbox
    {
        public static void AddTextboxToCallout(PowerPointSlide currentSlide, Selection selection)
        {
            if (!ShapeUtil.IsSelectionSingleShape(selection))
            {
                return;
            }
            AddTextboxToCallout(currentSlide, selection.ShapeRange[1]);
        }

        public static void AddTextboxToCallout(PowerPointSlide currentSlide, Shape callout)
        {
            Shape textbox = AddTextboxToSlide(currentSlide, callout.Left, callout.Top, callout.Width, callout.Height);
            Selection selection = AddTextboxActionHandler.GetNewSelection(callout, textbox);
            Shape group = selection.ShapeRange.Group();
            group.ZOrder(MsoZOrderCmd.msoSendBackward);
            textbox.Select(MsoTriState.msoTrue);
        }

        private static Shape AddTextboxToSlide(PowerPointSlide slide, float left, float top, float width, float height)
        {
            Shape textbox = slide.GetNativeSlide().Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            textbox.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
            textbox.ZOrder(MsoZOrderCmd.msoBringForward);
            return textbox;
        }

    }
}
