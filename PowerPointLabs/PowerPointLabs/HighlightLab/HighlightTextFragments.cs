using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.HighlightLab
{
    class HighlightTextFragments
    {
#pragma warning disable 0618
        public static Color backgroundColor = Color.FromArgb(255, 255, 0);
        public enum HighlightTextSelection { kShapeSelected, kTextSelected, kNoneSelected };
        public static HighlightTextSelection userSelection = HighlightTextSelection.kNoneSelected;
        public static void AddHighlightedTextFragments()
        {
            try
            {
                var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide as PowerPointSlide;

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
                        break;
                    case HighlightTextSelection.kNoneSelected:
                        return;
                    default:
                        return;
                }
                if (selectedText.Length <= 0)
                {
                    return;
                }

                if (selectedShapes.Count != 1)
                {
                    return;
                }

                List<PowerPoint.Shape> selectionToAnimate = GetShapesFromLinesInText(currentSlide, selectedText, selectedShapes[1]);
                GroupShapesForAnimation(selectionToAnimate);

                List<PowerPoint.Shape> shapesToAnimate = GetShapesToAnimate(currentSlide);
                SelectShapes(shapesToAnimate);

                RunAnimateInSlide();

            }
            catch (Exception e)
            {
                Logger.LogException(e, "AddHighlightedTextFragments");
                throw;
            }
        }

        private static void GroupShapesForAnimation(List<PowerPoint.Shape> shapes)
        {
            if (shapes.Count > 1)
            {
                Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
                
                foreach (PowerPoint.Shape sh in shapes)
                {
                    sh.Select(Office.MsoTriState.msoFalse);
                }

                PowerPoint.Shape groupedShape = 
                    Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Group();

                groupedShape.Name = "PPTLabsHighlightTextFragmentsShape" + Guid.NewGuid().ToString();
            }

            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
        }

        private static void SelectShapes(List<PowerPoint.Shape> shapes)
        {
            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();

            foreach (PowerPoint.Shape sh in shapes)
            {
                sh.Select(Office.MsoTriState.msoFalse);
            }
        }

        private static void RunAnimateInSlide()
        {
            bool oldFrameAnimationChecked = AnimateInSlide.frameAnimationChecked;
            AnimateInSlide.frameAnimationChecked = false;
            AnimateInSlide.isHighlightBullets = false;
            AnimateInSlide.isHighlightTextFragments = true;
            AnimateInSlide.AddAnimationInSlide();
            AnimateInSlide.isHighlightTextFragments = false;
            AnimateInSlide.frameAnimationChecked = oldFrameAnimationChecked;
        }

        private static List<PowerPoint.Shape> GetShapesToAnimate(PowerPointSlide currentSlide)
        {
            List<PowerPoint.Shape> previousFragments = currentSlide.GetTextFragments();
            currentSlide.RemoveAnimationsForShapes(previousFragments);

            return previousFragments;
        }

        private static List<PowerPoint.Shape> GetShapesFromLinesInText(PowerPointSlide currentSlide, Office.TextRange2 text, PowerPoint.Shape shape)
        {
            List<PowerPoint.Shape> shapesToAnimate = new List<PowerPoint.Shape>();
            Boolean isTextBoxTransparent = (shape.Fill.Transparency).CompareTo(1.0f) == 0;

            foreach (Office.TextRange2 line in text.Lines)
            {
                PowerPoint.Shape highlightShape = currentSlide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                    line.BoundLeft,
                    line.BoundTop,
                    line.BoundWidth,
                    line.BoundHeight);

                highlightShape.Adjustments[1] = 0.25f;
                highlightShape.Fill.ForeColor.RGB = Utils.Graphics.ConvertColorToRgb(backgroundColor);
                highlightShape.Fill.Transparency = 0.50f;
                highlightShape.Line.Visible = Office.MsoTriState.msoFalse;
                if (isTextBoxTransparent)
                {
                    Utils.Graphics.MoveZToJustBehind(highlightShape, shape);
                }
                highlightShape.Name = "PPTLabsHighlightTextFragmentsShape" + Guid.NewGuid().ToString();
                highlightShape.Tags.Add("HighlightTextFragment", highlightShape.Name);
                highlightShape.Select(Office.MsoTriState.msoFalse);
                shapesToAnimate.Add(highlightShape);
            }

            return shapesToAnimate;
        }
    }
}
