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
    class HighlightBulletsText
    {
        public static Color highlightColor = Color.FromArgb(242, 41, 10);
        public static Color defaultColor = Color.FromArgb(0, 0, 0);
        public enum HighlightTextSelection { kShapeSelected, kTextSelected, kNoneSelected };
        public static HighlightTextSelection userSelection = HighlightTextSelection.kNoneSelected;

        public static void AddHighlightBulletsText()
        {
            try
            {
                var currentSlide = PowerPointPresentation.CurrentSlide as PowerPointSlide;

                PowerPoint.ShapeRange selectedShapes = null;
                Office.TextRange2 selectedText = null;

                switch (userSelection)
                {
                    case HighlightTextSelection.kShapeSelected:
                        selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                        break;
                    case HighlightTextSelection.kTextSelected:
                        selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                        selectedText = Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange2.TrimText();
                        break;
                    case HighlightTextSelection.kNoneSelected:
                        currentSlide.DeleteShapesWithPrefix("PPTLabsIndicator");
                        currentSlide.DeleteShapesWithPrefix("PPTLabsHighlightBackgroundShape");
                        selectedShapes = currentSlide.Shapes.Range();
                        break;
                    default:
                        break;
                }

                List<PowerPoint.Shape> shapesToUse = GetShapesToUse(currentSlide, selectedShapes);
                if (currentSlide.Name.Contains("PPTLabsHighlightBulletsSlide"))
                    ProcessExistingHighlightSlide(currentSlide, shapesToUse);

                if (shapesToUse.Count == 0)
                    return;
                currentSlide.Name = "PPTLabsHighlightBulletsSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                PowerPoint.Sequence sequence = currentSlide.TimeLine.MainSequence;
                int initialEffectCount = sequence.Count;
                bool isFirstShape = IsFirstShape(currentSlide);

                foreach (PowerPoint.Shape sh in shapesToUse)
                {
                    if (!sh.Name.Contains("HighlightTextShape"))
                        sh.Name = "HighlightTextShape" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                    //Add Appear effect
                    sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor, PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                    int addedEffectCount = sequence.Count - initialEffectCount;
                    
                    //Add Disappear effect
                    sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor, PowerPoint.MsoAnimateByLevel.msoAnimateTextByFifthLevel, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    int addedEffectsStart = initialEffectCount + 1;
                    
                    RemoveEffectsForTextWithoutBullets(currentSlide, sh, addedEffectsStart, addedEffectCount, selectedText);
                    int finalEffectCount = sequence.Count - initialEffectCount;

                    if (finalEffectCount > 0)
                    {
                        FormatAddedEffects(currentSlide, addedEffectsStart, finalEffectCount, isFirstShape);                        
                        initialEffectCount += finalEffectCount;
                        isFirstShape = false;
                    }
                }
                AddAckSlide();
            }
            catch (Exception e)
            {
                //LogException(e, "SpotlightBtnClick");
                throw;
            }
        }

        private static void FormatAddedEffects(PowerPointSlide currentSlide, int addedEffectsStart, int finalEffectCount, bool isFirstShape)
        {
            PowerPoint.Sequence sequence = currentSlide.TimeLine.MainSequence;

            //Highlight Color appear
            PowerPoint.Effect firstHighlightAppear = sequence[addedEffectsStart];
            firstHighlightAppear.EffectParameters.Color2.RGB = CreateRGB(highlightColor);
            firstHighlightAppear.Timing.Duration = 0.01f;
            firstHighlightAppear.Timing.TriggerType = isFirstShape ? PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick : PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious;

            int countCopy = finalEffectCount / 2;
            for (int i = 2, j = 1; i < finalEffectCount; i += 2, j++)
            {
                PowerPoint.Effect nextHighlightAppear = sequence[addedEffectsStart - 1 + i];
                nextHighlightAppear.EffectParameters.Color2.RGB = CreateRGB(highlightColor);
                nextHighlightAppear.Timing.Duration = 0.01f;

                PowerPoint.Effect firstHighlightDisappear = sequence[addedEffectsStart - 1 + countCopy + j];
                firstHighlightDisappear.EffectParameters.Color2.RGB = CreateRGB(defaultColor);
                firstHighlightDisappear.Timing.Duration = 0.01f;
                firstHighlightDisappear.MoveTo(addedEffectsStart + i);
                firstHighlightDisappear.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
            }

            PowerPoint.Effect lastHighlightDisappear = sequence[sequence.Count];
            lastHighlightDisappear.EffectParameters.Color2.RGB = CreateRGB(defaultColor);
            lastHighlightDisappear.Timing.Duration = 0.01f;
            lastHighlightDisappear.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
        }

        private static void ProcessExistingHighlightSlide(PowerPointSlide currentSlide, List<PowerPoint.Shape> shapesToUse)
        {
            currentSlide.DeleteShapesWithPrefix("PPTLabsIndicator");
            currentSlide.DeleteShapesWithPrefix("PPTLabsHighlightBackgroundShape");

            foreach (PowerPoint.Shape tmp in currentSlide.Shapes)
                if (shapesToUse.Contains(tmp))
                    if (userSelection != HighlightTextSelection.kTextSelected)
                        currentSlide.DeleteShapeAnimations(tmp);
        }

        private static void RemoveEffectsForTextWithoutBullets(PowerPointSlide currentSlide, PowerPoint.Shape sh, int addedEffectsStart, int addedEffectCount, Office.TextRange2 selectedText)
        {
            PowerPoint.Sequence sequence = currentSlide.TimeLine.MainSequence;
            //Remove effects for text without bullets
            for (int i = 1, j = 1; i <= sh.TextFrame2.TextRange.Paragraphs.Count; i++, j++)
            {
                Office.TextRange2 paragraph = sh.TextFrame2.TextRange.Paragraphs[i];
                if (userSelection == HighlightTextSelection.kTextSelected)
                {
                    if (paragraph.Text.Trim().Length == 0)
                    {
                        addedEffectCount--;
                        j--;
                        continue;
                    }
                    if ((selectedText.Start + selectedText.Length < paragraph.Start) || (selectedText.Start > paragraph.Start + paragraph.Length - 1))
                    {
                        sequence[addedEffectsStart - 1 + i + addedEffectCount].Delete();
                        sequence[addedEffectsStart - 1 + j].Delete();
                        j--;
                        addedEffectCount -= 2;
                    }
                }
                else
                {
                    if (paragraph.ParagraphFormat.Bullet.Visible == Office.MsoTriState.msoFalse)
                    {
                        sequence[addedEffectsStart - 1 + i + addedEffectCount].Delete(); //Delete disappear effect
                        sequence[addedEffectsStart - 1 + j].Delete(); //Delete appear effect
                        j--;
                        addedEffectCount -= 2;
                    }
                }
            }
        }
        private static bool IsFirstShape(PowerPointSlide currentSlide)
        {
            PowerPoint.Sequence sequence = currentSlide.TimeLine.MainSequence;
            bool isFirstShape = true;
            if (sequence.Count != 0)
            {
                isFirstShape = (sequence[sequence.Count].EffectType == PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor) ? false : true;
                if (sequence[1].EffectType == PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontColor)
                    sequence[1].Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
            }
            return isFirstShape;
        }

        private static List<PowerPoint.Shape> GetShapesToUse(PowerPointSlide currentSlide, PowerPoint.ShapeRange selectedShapes)
        {
            List<PowerPoint.Shape> shapesToUse = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape sh in selectedShapes)
            {
                if (userSelection != HighlightTextSelection.kTextSelected)
                {
                    if (sh.HasTextFrame == Office.MsoTriState.msoTrue && sh.TextFrame2.HasText == Office.MsoTriState.msoTrue
                    && sh.TextFrame2.TextRange.ParagraphFormat.Bullet.Visible == Office.MsoTriState.msoTrue
                    && sh.TextFrame2.TextRange.ParagraphFormat.Bullet.Type != Office.MsoBulletType.msoBulletNone)
                    {
                        shapesToUse.Add(sh);
                    }
                }
                else
                {
                    if (sh.HasTextFrame == Office.MsoTriState.msoTrue && sh.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                    {
                        shapesToUse.Add(sh);
                    }
                }
            }
            return shapesToUse;
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
