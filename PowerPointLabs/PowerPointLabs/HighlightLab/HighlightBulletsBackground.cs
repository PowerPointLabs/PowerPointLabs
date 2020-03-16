using System;
using System.Collections.Generic;
using System.Linq;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.AnimationLab;
using PowerPointLabs.Models;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.HighlightLab
{
    class HighlightBulletsBackground
    {
#pragma warning disable 0618
        public enum HighlightBackgroundSelection { kShapeSelected, kTextSelected, kNoneSelected };
        public static HighlightBackgroundSelection userSelection = HighlightBackgroundSelection.kNoneSelected;
        public static bool IsHighlightBackgroundEnabled { get; set; } = true;

        public static void AddHighlightBulletsBackground()
        {
            try
            {
                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide as PowerPointSlide;
                currentSlide.Name = "PPTLabsHighlightBulletsSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                PowerPoint.ShapeRange selectedShapes = null;
                Office.TextRange2 selectedText = null;

                //Get shapes to consider for animation
                List<PowerPoint.Shape> shapesToUse = null;
                switch (userSelection)
                {
                    case HighlightBackgroundSelection.kShapeSelected:
                        selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                        shapesToUse = GetShapesToUse(currentSlide, selectedShapes);
                        break;
                    case HighlightBackgroundSelection.kTextSelected:
                        selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                        selectedText = Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange2.TrimText();
                        shapesToUse = GetShapesToUse(currentSlide, selectedShapes);
                        break;
                    case HighlightBackgroundSelection.kNoneSelected:
                        currentSlide.DeleteIndicator();
                        shapesToUse = GetAllUsableShapesInSlide(currentSlide);
                        break;
                    default:
                        break;
                }

                Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(currentSlide.Index);

                if (shapesToUse == null || shapesToUse.Count == 0)
                {
                    return;
                }

                SelectOldShapesToAnimate(currentSlide, shapesToUse);
                bool newShapesAdded = AddNewShapesToAnimate(currentSlide, shapesToUse, selectedText);

                if (newShapesAdded)
                {
                    bool oldValue = AnimationLabSettings.IsUseFrameAnimation;
                    AnimationLabSettings.IsUseFrameAnimation = false;
                    AnimateInSlide.AddAnimationInSlide(isHighlightBullets: true);
                    AnimationLabSettings.IsUseFrameAnimation = oldValue;
                    PowerPointPresentation.Current.AddAckSlide();
                }

                if (currentSlide.HasAnimationForClick(clickNumber: 1))
                {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
                }

                Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AddHighlightBulletsBackground");
                throw;
            }
        }

        //Add highlight shape for paragraphs within selected shape which have bullets or with text selected by user
        private static bool AddNewShapesToAnimate(PowerPointSlide currentSlide, List<PowerPoint.Shape> shapesToUse, Office.TextRange2 selectedText)
        {
            bool anySelected = false;

            foreach (PowerPoint.Shape sh in shapesToUse)
            {
                sh.Name = "HighlightBackgroundShape" + Guid.NewGuid();
            }

            if (userSelection == HighlightBackgroundSelection.kTextSelected)
            {
                foreach (PowerPoint.Shape sh in shapesToUse)
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
                foreach (PowerPoint.Shape sh in shapesToUse)
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
            PowerPoint.Shape highlightShape = currentSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                                                            paragraph.BoundLeft,
                                                            paragraph.BoundTop,
                                                            paragraph.BoundWidth,
                                                            paragraph.BoundHeight);
            highlightShape.Adjustments[1] = 0.25f;
            highlightShape.Fill.ForeColor.RGB = Utils.GraphicsUtil.ConvertColorToRgb(HighlightLabSettings.bulletsBackgroundColor);
            highlightShape.Fill.Transparency = 0.50f;
            highlightShape.Line.Visible = Office.MsoTriState.msoFalse;
            Utils.ShapeUtil.MoveZToJustBehind(highlightShape, sh);
            highlightShape.Name = "PPTLabsHighlightBackgroundShape" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
            highlightShape.Tags.Add("HighlightBackground", sh.Name);
            highlightShape.Select(Office.MsoTriState.msoFalse);
        }

        private static void SelectOldShapesToAnimate(PowerPointSlide currentSlide, List<PowerPoint.Shape> shapesToUse)
        {
            List<PowerPoint.Shape> shapesToDelete = new List<PowerPoint.Shape>();
            bool shouldSelect;

            List<PowerPoint.Shape> shapes = currentSlide.GetShapesOrderedByTimeline();
            foreach (PowerPoint.Shape sh in shapes)
            {
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
            }

            if (shapesToDelete.Count > 0)
            {
                foreach (PowerPoint.Shape sh in shapesToDelete)
                {
                    sh.SafeDelete();
                }
            }
        }

        /// <summary>
        /// Get shapes to use for animation.
        /// If user does not select anything: Select shapes which have bullet points
        /// If user selects some shapes: Keep shapes from user selection which have bullet points
        /// If user selects some text: Keep shapes used to store text
        /// </summary>
        private static List<PowerPoint.Shape> GetShapesToUse(PowerPointSlide currentSlide, PowerPoint.ShapeRange shapes)
        {
            List<PowerPoint.Shape> shapesToUse = shapes.Cast<PowerPoint.Shape>()
                                    .Where(sh => !sh.Name.Contains("PPTLabsHighlightBackgroundShape")
                                                    && HasText(sh))
                                    .ToList();
            return shapesToUse;
        }

        /// <summary>
        /// Get all shapes in slide to use for animation.
        /// If there are text boxes with bullet points, returns only the text boxes with bullet points.
        /// If there are no text boxes with bullet points, returns everything.
        /// </summary>
        private static List<PowerPoint.Shape> GetAllUsableShapesInSlide(PowerPointSlide currentSlide)
        {
            PowerPoint.Shape[] selectedShapes = currentSlide.Shapes.Range().Cast<PowerPoint.Shape>().ToArray();

            List<PowerPoint.Shape> usableShapesWithBullets = selectedShapes
                                            .Where(sh => !sh.Name.Contains("PPTLabsHighlightBackgroundShape")
                                                        && HasText(sh)
                                                        && HasBullets(sh))
                                            .ToList();

            List<PowerPoint.Shape> allUsableShapes = selectedShapes
                                    .Where(sh => !sh.Name.Contains("PPTLabsHighlightBackgroundShape")
                                                    && HasText(sh))
                                    .ToList();

            if (usableShapesWithBullets.Count == 0)
            {
                return allUsableShapes;
            }
            else
            {
                return usableShapesWithBullets;
            }
        }

        /// <summary>
        /// Returns true iff shape (assuming has text) has bullet points.
        /// Duplicate method in HighlightBulletsText.cs
        /// </summary>
        private static bool HasBullets(PowerPoint.Shape shape)
        {
            return shape.TextFrame2.TextRange.ParagraphFormat.Bullet.Visible == Office.MsoTriState.msoTrue &&
                   shape.TextFrame2.TextRange.ParagraphFormat.Bullet.Type != Office.MsoBulletType.msoBulletNone;

        }

        /// <summary>
        /// Returns true iff shape has a text frame.
        /// Duplicate method in HighlightBulletsText.cs
        /// </summary>
        private static bool HasText(PowerPoint.Shape shape)
        {
            return shape.HasTextFrame == Office.MsoTriState.msoTrue &&
                   shape.TextFrame2.HasText == Office.MsoTriState.msoTrue;

        }
    }
}
