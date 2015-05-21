using System;
using System.Collections.Generic;
using System.Drawing;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    class HighlightBulletsBackground
    {
        public static Color backgroundColor = Color.FromArgb(255, 255, 0);
        public enum HighlightBackgroundSelection { kShapeSelected, kTextSelected, kNoneSelected };
        public static HighlightBackgroundSelection userSelection = HighlightBackgroundSelection.kNoneSelected;
        public static void AddHighlightBulletsBackground()
        {
            try
            {
                var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide as PowerPointSlide;
                currentSlide.Name = "PPTLabsHighlightBulletsSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                PowerPoint.ShapeRange selectedShapes = null;
                Office.TextRange2 selectedText = null;

                //Get shapes to consider for animation
                switch (userSelection)
                {
                    case HighlightBackgroundSelection.kShapeSelected:
                        selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                        break;
                    case HighlightBackgroundSelection.kTextSelected:
                        selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                        selectedText = Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange2.TrimText();
                        break;
                    case HighlightBackgroundSelection.kNoneSelected:
                        currentSlide.DeleteShapesWithPrefix("PPIndicator");
                        currentSlide.DeleteShapesWithPrefix("PPTLabsHighlightBackgroundShape");
                        selectedShapes = currentSlide.Shapes.Range();
                        break;
                    default:
                        break;
                }

                List<PowerPoint.Shape> shapesToUse = GetShapesToUse(currentSlide, selectedShapes);
                
                bool newShapesAdded = false;
                Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(currentSlide.Index);

                SelectOldShapesToAnimate(currentSlide, shapesToUse);
                newShapesAdded = AddNewShapesToAnimate(currentSlide, shapesToUse, selectedText);

                if (newShapesAdded)
                {
                    bool oldValue = AnimateInSlide.frameAnimationChecked;
                    AnimateInSlide.frameAnimationChecked = false;
                    AnimateInSlide.isHighlightBullets = true;
                    AnimateInSlide.AddAnimationInSlide();
                    AnimateInSlide.frameAnimationChecked = oldValue;
                    PowerPointPresentation.Current.AddAckSlide();
                }
                Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddHighlightBulletsBackground");
                throw;
            }
        }

        //Add highlight shape for paragraphs within selected shape which have bullets or with text selected by user
        private static bool AddNewShapesToAnimate(PowerPointSlide currentSlide, List<PowerPoint.Shape> shapesToUse, Office.TextRange2 selectedText)
        {
            bool anySelected = false;

            foreach (var sh in shapesToUse)
            {
                sh.Name = "HighlightBackgroundShape" + Guid.NewGuid();
            }

            if (userSelection == HighlightBackgroundSelection.kTextSelected)
            {
                foreach (var sh in shapesToUse)
                {
                    foreach (Office.TextRange2 paragraph in sh.TextFrame2.TextRange.Paragraphs)
                    {
                        if (paragraph.Start <= selectedText.Start + selectedText.Length
                            && selectedText.Start <= paragraph.Start + paragraph.Length - 1
                            && paragraph.TrimText().Length > 0)
                        {
                            GenerateHighlightShape(currentSlide, paragraph, sh);
                            anySelected = true;
                        }
                    }
                }
            }
            else
            {
                foreach (var sh in shapesToUse)
                {
                    bool anySelectedForShape = false;
                    foreach (Office.TextRange2 paragraph in sh.TextFrame2.TextRange.Paragraphs)
                    {
                        if (paragraph.ParagraphFormat.Bullet.Visible == Office.MsoTriState.msoTrue
                            && paragraph.TrimText().Length > 0)
                        {
                            GenerateHighlightShape(currentSlide, paragraph, sh);
                            anySelected = true;
                            anySelectedForShape = true;
                        }
                    }
                    if (anySelectedForShape)
                    {
                        continue;
                    }
                    foreach (Office.TextRange2 paragraph in sh.TextFrame2.TextRange.Paragraphs)
                    {
                        if (paragraph.TrimText().Length > 0)
                        {
                            GenerateHighlightShape(currentSlide, paragraph, sh);
                            anySelected = true;
                        }
                    }
                }
            }
            return anySelected;
        }

        private static void GenerateHighlightShape(PowerPointSlide currentSlide, Office.TextRange2 paragraph, PowerPoint.Shape sh)
        {
            var highlightShape = currentSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                                                            paragraph.BoundLeft,
                                                            paragraph.BoundTop,
                                                            paragraph.BoundWidth,
                                                            paragraph.BoundHeight);
            highlightShape.Adjustments[1] = 0.25f;
            highlightShape.Fill.ForeColor.RGB = Utils.Graphics.ConvertColorToRgb(backgroundColor);
            highlightShape.Fill.Transparency = 0.50f;
            highlightShape.Line.Visible = Office.MsoTriState.msoFalse;
            Utils.Graphics.MoveZToJustBehind(highlightShape, sh);
            highlightShape.Name = "PPTLabsHighlightBackgroundShape" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
            highlightShape.Tags.Add("HighlightBackground", sh.Name);
            highlightShape.Select(Office.MsoTriState.msoFalse);
        }

        private static void SelectOldShapesToAnimate(PowerPointSlide currentSlide, List<PowerPoint.Shape> shapesToUse)
        {
            List<PowerPoint.Shape> shapesToDelete = new List<PowerPoint.Shape>();
            bool shouldSelect;

            for (int i = currentSlide.Shapes.Count; i >= 1; i--)
            {
                PowerPoint.Shape sh = currentSlide.Shapes[i];
                shouldSelect = true; //We should not select existing highlight shapes. Instead they should be deleted
                if (sh.Name.Contains("PPTLabsHighlightBackgroundShape"))
                {
                    if (userSelection != HighlightBackgroundSelection.kTextSelected)
                    {
                        foreach (PowerPoint.Shape tmp in shapesToUse)
                        {
                            //Each highlight shape stores a tag of the shape it is associated with
                            if (sh.Tags["HighlightBackground"].Equals(tmp.Name))
                            {
                                shapesToDelete.Add(sh);
                                shouldSelect = false;
                                break;
                            }
                        }
                    }
                    if (shouldSelect)
                    {
                        currentSlide.DeleteShapeAnimations(sh);
                        sh.Select(Office.MsoTriState.msoFalse);
                    }
                }
                //Remove existing animations for highlight text as well
                if (sh.Name.Contains("HighlightTextShape"))
                    currentSlide.DeleteShapeAnimations(sh);
            }

            if (shapesToDelete.Count > 0)
                foreach (PowerPoint.Shape sh in shapesToDelete)
                    sh.Delete();
        }

        /// <summary>
        /// Get shapes to use for animation.
        /// If user does not select anything: Select shapes which have bullet points
        /// If user selects some shapes: Keep shapes from user selection which have bullet points
        /// If user selects some text: Keep shapes used to store text
        /// </summary>
        private static List<PowerPoint.Shape> GetShapesToUse(PowerPointSlide currentSlide, PowerPoint.ShapeRange shapes)
        {
            List<PowerPoint.Shape> shapesToUse = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape sh in shapes)
            {
                if (sh.Name.Contains("PPTLabsHighlightBackgroundShape"))
                    continue;
                if (sh.HasTextFrame == Office.MsoTriState.msoTrue && sh.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                {
                    currentSlide.DeleteShapeAnimations(sh);
                    shapesToUse.Add(sh);
                }
            }
            return shapesToUse;
        }
    }
}
