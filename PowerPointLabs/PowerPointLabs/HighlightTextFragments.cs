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
                        break;
                    case HighlightTextSelection.kNoneSelected:
                        return;
                    default:
                        break;
                }

                List<PowerPoint.Shape> selectionToAnimate = GetShapesFromLinesInText(currentSlide, selectedText);

                List<PowerPoint.Shape> shapesToAnimate = GetShapesToAnimate(currentSlide, selectionToAnimate);

                AddAnimationForShapes(shapesToAnimate, currentSlide);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddHighlightedTextFragments");
                throw;
            }
        }

        private static List<PowerPoint.Shape> GetShapesToAnimate(PowerPointSlide currentSlide,
            List<PowerPoint.Shape> selectionToAnimate)
        {
            List<PowerPoint.Shape> previousFragments = currentSlide.getTextFragments();
            currentSlide.RemoveAnimationsForShapes(previousFragments);

            previousFragments.Reverse();

            return previousFragments;
        }

        private static List<PowerPoint.Shape> GetShapesFromLinesInText(PowerPointSlide currentSlide, Office.TextRange2 text)
        {
            List<PowerPoint.Shape> shapesToAnimate = new List<PowerPoint.Shape>();

            foreach (Office.TextRange2 line in text.Lines)
            {
                PowerPoint.Shape highlightShape = currentSlide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                    line.BoundLeft,
                    line.BoundTop,
                    line.BoundWidth,
                    line.BoundHeight);

                highlightShape.Adjustments[1] = 0.25f;
                highlightShape.Fill.ForeColor.RGB = PowerPointLabsGlobals.CreateRGB(backgroundColor);
                highlightShape.Fill.Transparency = 0.50f;
                highlightShape.Line.Visible = Office.MsoTriState.msoFalse;
                highlightShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                highlightShape.Name = "PPTLabsHighlightTextFragmentsShape" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                highlightShape.Tags.Add("HighlightTextFragment", highlightShape.Name);
                highlightShape.Select(Office.MsoTriState.msoFalse);
                shapesToAnimate.Add(highlightShape);
            }

            return shapesToAnimate;
        }

        private static void AddAnimationForShapes(List<PowerPoint.Shape> shapesToAnimate,
            PowerPointSlide currentSlide)
        {
            for (int num = 0; num < shapesToAnimate.Count - 1; num++)
            {
                PowerPoint.Shape shape1 = shapesToAnimate[num];
                PowerPoint.Shape shape2 = shapesToAnimate[num + 1];

                if (shape1 == null || shape2 == null)
                    return;

                if (num == 0)
                {
                    PowerPoint.Effect appear = currentSlide.TimeLine.MainSequence.AddEffect(
                        shape1,
                        PowerPoint.MsoAnimEffect.msoAnimEffectAppear,
                        PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                }

                //if (NeedsFrameAnimation(shape1, shape2))
                //{
                //    FrameMotionAnimation.animationType = FrameMotionAnimation.FrameMotionAnimationType.kInSlideAnimate;
                //    FrameMotionAnimation.AddFrameMotionAnimation(currentSlide, shape1, shape2, 0.5f);
                //}
                //else
                    DefaultMotionAnimation.AddDefaultMotionAnimation(currentSlide,
                        shape1,
                        shape2,
                        0.5f,
                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);

                //Transition from shape1 to shape2
                PowerPoint.Effect shape2Appear = currentSlide.TimeLine.MainSequence.AddEffect(
                    shape2,
                    PowerPoint.MsoAnimEffect.msoAnimEffectAppear,
                    PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                PowerPoint.Effect shape1Disappear = currentSlide.TimeLine.MainSequence.AddEffect(
                    shape1,
                    PowerPoint.MsoAnimEffect.msoAnimEffectAppear,
                    PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                shape1Disappear.Exit = Office.MsoTriState.msoTrue;
            }
        }

        private static void log(string s)
        {
            System.Diagnostics.Debug.WriteLine(s);
        }
    }
}
