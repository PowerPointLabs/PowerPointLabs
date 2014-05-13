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
    class HighlightBulletsBackground
    {
        public static Color backgroundColor = Color.FromArgb(255, 255, 0);
        public enum HighlightBulletsSelection { kShapeSelected, kTextSelected, kNoneSelected };
        public static HighlightBulletsSelection userSelection = HighlightBulletsSelection.kNoneSelected;
        public static void AddHighlightBulletsBackground()
        {
            try
            {
                var currentSlide = PowerPointPresentation.CurrentSlide as PowerPointSlide;
                currentSlide.Name = "PPTLabsHighlightBulletsSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                PowerPoint.ShapeRange selectedShapes = null;
                Office.TextRange2 selectedText = null;

                switch (userSelection)
                {
                    case HighlightBulletsSelection.kShapeSelected:
                        selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                        break;
                    case HighlightBulletsSelection.kTextSelected:
                        selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                        selectedText = Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange2.TrimText();
                        break;
                    case HighlightBulletsSelection.kNoneSelected:
                        currentSlide.DeleteShapesWithPrefix("PPTLabsIndicator");
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

                SelectShapesToAnimate(currentSlide, shapesToUse);
                newShapesAdded = AddNewShapesToAnimate(currentSlide, shapesToUse, selectedText);

                if (newShapesAdded)
                {
                    bool oldValue = AnimateInSlide.frameAnimationChecked;
                    AnimateInSlide.frameAnimationChecked = false;
                    AnimateInSlide.isHighlightBullets = true;
                    AnimateInSlide.AddAnimationInSlide();
                    AnimateInSlide.frameAnimationChecked = oldValue;
                    AddAckSlide();
                }
                Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            }
            catch (Exception e)
            {
                //LogException(e, "SpotlightBtnClick");
                throw;
            }
        }

        private static bool AddNewShapesToAnimate(PowerPointSlide currentSlide, List<PowerPoint.Shape> shapesToUse, Office.TextRange2 selectedText)
        {
            bool anySelected = false;
            foreach (PowerPoint.Shape sh in shapesToUse)
            {
                if (!sh.Name.Contains("HighlightBackgroundShape"))
                    sh.Name = "HighlightBackgroundShape" + Guid.NewGuid().ToString();
                foreach (Office.TextRange2 paragraph in sh.TextFrame2.TextRange.Paragraphs)
                {
                    if ((userSelection != HighlightBulletsSelection.kTextSelected) && (paragraph.ParagraphFormat.Bullet.Visible == Office.MsoTriState.msoTrue && paragraph.TrimText().Length > 0)
                        || ((userSelection == HighlightBulletsSelection.kTextSelected) && (!((selectedText.Start + selectedText.Length < paragraph.Start) || (selectedText.Start > paragraph.Start + paragraph.Length - 1)) && paragraph.TrimText().Length > 0)))
                    {
                        PowerPoint.Shape highlightShape = currentSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, paragraph.BoundLeft, paragraph.BoundTop, paragraph.BoundWidth, paragraph.BoundHeight);
                        highlightShape.Adjustments[1] = 0.25f;
                        highlightShape.Fill.ForeColor.RGB = CreateRGB(backgroundColor);
                        highlightShape.Fill.Transparency = 0.50f;
                        highlightShape.Line.Visible = Office.MsoTriState.msoFalse;
                        highlightShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                        highlightShape.Name = "PPTLabsHighlightBackgroundShape" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                        highlightShape.Tags.Add("HighlightBackground", sh.Name);
                        highlightShape.Select(Office.MsoTriState.msoFalse);
                        anySelected = true;
                    }
                }
            }
            return anySelected;
        }

        private static void SelectShapesToAnimate(PowerPointSlide currentSlide, List<PowerPoint.Shape> shapesToUse)
        {
            List<PowerPoint.Shape> shapesToDelete = new List<PowerPoint.Shape>();
            bool shouldSelect;

            for (int i = currentSlide.Shapes.Count; i >= 1; i--)
            {
                PowerPoint.Shape sh = currentSlide.Shapes[i];
                shouldSelect = true;
                if (sh.Name.Contains("PPTLabsHighlightBackgroundShape"))
                {
                    if (userSelection != HighlightBulletsSelection.kTextSelected)
                    {
                        foreach (PowerPoint.Shape tmp in shapesToUse)
                        {
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
                if (sh.Name.Contains("HighlightTextShape"))
                    currentSlide.DeleteShapeAnimations(sh);
            }

            if (shapesToDelete.Count > 0)
                foreach (PowerPoint.Shape sh in shapesToDelete)
                    sh.Delete();
        }

        private static int CreateRGB(Color color)
        {
            // initial value
            int rgb = 0;

            // swap
            int red = color.B;
            int blue = color.R;
            int green = color.G;

            // create the newColor
            Color newColor = Color.FromArgb(red, green, blue);

            // set the return value
            rgb = newColor.ToArgb();

            // return value
            return rgb;
        }

        private static List<PowerPoint.Shape> GetShapesToUse(PowerPointSlide currentSlide, PowerPoint.ShapeRange shapes)
        {
            List<PowerPoint.Shape> shapesToUse = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape sh in shapes)
            {
                if (sh.Name.Contains("PPTLabsHighlightBackgroundShape"))
                    continue;
                if (userSelection != HighlightBulletsSelection.kTextSelected)
                {
                    if (sh.HasTextFrame == Office.MsoTriState.msoTrue && sh.TextFrame2.HasText == Office.MsoTriState.msoTrue
                    && sh.TextFrame2.TextRange.ParagraphFormat.Bullet.Visible == Office.MsoTriState.msoTrue
                    && sh.TextFrame2.TextRange.ParagraphFormat.Bullet.Type != Office.MsoBulletType.msoBulletNone)
                    {
                        currentSlide.DeleteShapeAnimations(sh);    
                        shapesToUse.Add(sh);
                    }
                }
                else
                {
                    if (sh.HasTextFrame == Office.MsoTriState.msoTrue && sh.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                    {
                        currentSlide.DeleteShapeAnimations(sh);
                        shapesToUse.Add(sh);
                    }
                }
            }
            return shapesToUse;
        }

        private static void AddAckSlide()
        {
            try
            {
                PowerPointSlide lastSlide = PowerPointPresentation.Slides.Last();
                if (!lastSlide.isAckSlide())
                {
                    lastSlide.CreateAckSlide();
                }
            }
            catch (Exception e)
            {
                //LogException(e, "AddAckSlide");
                throw;
            }
        }
    }
}
