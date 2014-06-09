using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    class HighlightTextFragments
    {
        public static Color backgroundColor = Color.FromArgb(255, 255, 0);
        public enum HighlightTextSelection { kShapeSelected, kTextSelected, kNoneSelected };
        public static HighlightTextSelection userSelection = HighlightTextSelection.kNoneSelected;
        public static void AddHighlightedTextFragments()
        {
            try
            {
                var currentSlide = PowerPointPresentation.CurrentSlide as PowerPointSlide;

                PowerPoint.ShapeRange selectedShapes = null;
                Office.TextRange2 selectedText = null;

                //Get shapes to consider for animation
                switch (userSelection)
                {
                    case HighlightTextSelection.kShapeSelected:
                        return;
                    case HighlightTextSelection.kTextSelected:
                        selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                        selectedText = Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange2.TrimText();

                        System.Diagnostics.Debug.WriteLine(selectedText);
                        break;
                    case HighlightTextSelection.kNoneSelected:
                        return;
                    default:
                        break;
                }

                
                var selectedTextRange = Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange;

                PowerPoint.TextFrame2 currentTextFrame = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1].TextFrame2;

                float topOffset = currentTextFrame.MarginTop;

                Office.TextRange2 firstParagraph = currentTextFrame.TextRange.Paragraphs[1];

                if (selectedTextRange.Start >= firstParagraph.Start &&
                    selectedTextRange.Start < (firstParagraph.Start + firstParagraph.Text.Length))
                {
                    topOffset = -1 * (selectedTextRange.ParagraphFormat.SpaceBefore);
                    topOffset += currentTextFrame.MarginTop;
                }

                PowerPoint.Shape textBox = PowerPointPresentation.CurrentSlide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal, selectedTextRange.BoundLeft,
                    selectedTextRange.BoundTop + topOffset,
                    selectedTextRange.BoundWidth, selectedTextRange.BoundHeight);

                textBox.TextFrame.TextRange.Font.Size = selectedTextRange.Font.Size;
                textBox.TextFrame.TextRange.Font.Color.RGB = selectedTextRange.Font.Color.RGB;

                textBox.TextFrame2.MarginLeft = 0;
                textBox.TextFrame2.MarginRight = 0;

                textBox.TextFrame2.WordWrap = Office.MsoTriState.msoFalse;
                textBox.TextFrame2.TextRange.ParagraphFormat.Bullet.Visible = Office.MsoTriState.msoFalse;

                textBox.TextFrame2.TextRange.ParagraphFormat.SpaceBefore = 0;
                textBox.TextFrame2.TextRange.ParagraphFormat.SpaceAfter = 0;

                textBox.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeShapeToFitText;

                textBox.TextFrame.TextRange.InsertAfter(selectedTextRange.Text);
               
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddHighlightedTextFragments");
                throw;
            }
        }
    }
}
