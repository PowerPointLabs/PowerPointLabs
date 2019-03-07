using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
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

        public static Shape AddTextboxToCallout(PowerPointSlide currentSlide, Shape callout)
        {
            Shape textbox = AddTextboxToSlide(currentSlide, callout.Left, callout.Top, callout.Width, callout.Height);
            string[] rangeArray = new string[2];
            rangeArray[0] = callout.Name;
            rangeArray[1] = textbox.Name;
            Shape group = currentSlide.Shapes.Range(rangeArray).Group();
            group.ZOrder(MsoZOrderCmd.msoSendBackward);
            textbox.Select(MsoTriState.msoTrue);
            return group;
        }

        /// <summary>
        /// Text alignment is default to be centralized
        /// </summary>
        private static Shape AddTextboxToSlide(PowerPointSlide slide, float left, float top, float width, float height)
        {
            //Create textbox with the correct size, position and text alignment
            Slide nativeSlide = slide.GetNativeSlide();
            Shape textbox = TextBoxWithDefaultText.CreateTextBox(nativeSlide, MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            textbox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            textbox.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;

            //Set placeholder text
            textbox.AlternativeText = "Enter text here.";

            textbox.ZOrder(MsoZOrderCmd.msoBringForward);

            //Somehow without explicitly restating the variables, the textbox size will be wrong.
            textbox.Left = left;
            textbox.Top = top;
            textbox.Width = width;
            textbox.Height = height;

            return textbox;
        }

    }
}
