using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    class InSlideAnimate
    {
        public static void AddInSlideAnimation()
        {
            try
            {
                //Get References of current and next slides
                var currentSlide = PowerPointPresentation.CurrentSlide;
                PowerPoint.ShapeRange shapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

                foreach (PowerPoint.Shape sh in shapes)
                {
                    currentSlide.RemoveAnimationsForShape(sh);
                }

                PowerPoint.Sequence sequence = currentSlide.TimeLine.MainSequence;
                PowerPoint.Effect effectMotion = null;
                PowerPoint.Effect effectResize = null;
                PowerPoint.Effect effectRotate = null;
                PowerPoint.MsoAnimTriggerType trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;

                for (int num = 1; num <= shapes.Count - 1; num++)
                {
                    PowerPoint.Shape sh1 = shapes[num];
                    PowerPoint.Shape sh2 = shapes[num + 1];

                    if (sh1 == null || sh2 == null)
                        return;

                    if (num == 1)
                    {
                        PowerPoint.Effect appear = sequence.AddEffect(sh1, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                    }

                    trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                    float finalX = (sh2.Left + (sh2.Width) / 2);
                    float initialX = (sh1.Left + (sh1.Width) / 2);
                    float finalY = (sh2.Top + (sh2.Height) / 2);
                    float initialY = (sh1.Top + (sh1.Height) / 2);

                    float finalRotation = sh2.Rotation;
                    float initialRotation = sh1.Rotation;

                    float finalWidth = sh2.Width;
                    float initialWidth = sh1.Width;
                    float finalHeight = sh2.Height;
                    float initialHeight = sh1.Height;
                    float finalFont = 0.0f;
                    float initialFont = 0.0f;
                    int numFrames = (int)(defaultDuration / 0.04f);
                    numFrames = (numFrames > 30) ? 30 : numFrames;

                    if (sh1.HasTextFrame == Office.MsoTriState.msoTrue && (sh1.TextFrame.HasText == Office.MsoTriState.msoTriStateMixed || sh1.TextFrame.HasText == Office.MsoTriState.msoTrue) && sh1.TextFrame.TextRange.Font.Size != sh2.TextFrame.TextRange.Font.Size)
                    {
                        finalFont = sh2.TextFrame.TextRange.Font.Size;
                        initialFont = sh1.TextFrame.TextRange.Font.Size;
                    }

                    if ((frameAnimationChecked && (finalHeight != initialHeight || finalWidth != initialWidth))
                        || ((initialRotation != finalRotation || initialRotation % 90 != 0) && (finalHeight != initialHeight || finalWidth != initialWidth))
                        || finalFont != initialFont)
                    {
                        float incrementWidth = ((finalWidth / initialWidth) - 1.0f) / numFrames;
                        float incrementHeight = ((finalHeight / initialHeight) - 1.0f) / numFrames;
                        float incrementRotation = GetMinimumRotation(initialRotation, finalRotation) / numFrames;
                        float incrementLeft = (finalX - initialX) / numFrames;
                        float incrementTop = (finalY - initialY) / numFrames;
                        float incrementFont = (finalFont - initialFont) / numFrames;

                        //PowerPoint.Effect shapeEffect = GetShapeAnnimations(addedSlide, sh1);
                        //if (shapeEffect != null)
                        //    shapeEffect.Delete();

                        PowerPoint.Shape lastShape = sh1;
                        for (int i = 1; i <= numFrames; i++)
                        {
                            PowerPoint.Shape dupShape = sh1.Duplicate()[1];
                            if (i != 1)
                            {
                                sequence[sequence.Count].Delete();
                            }
                            PowerPoint.Effect shapeEffect = GetShapeAnnimations(currentSlide, dupShape);
                            if (shapeEffect != null)
                                shapeEffect.Delete();

                            dupShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                            dupShape.Left = sh1.Left;
                            dupShape.Top = sh1.Top;

                            if (incrementWidth != 0.0f)
                            {
                                dupShape.ScaleWidth((1.0f + (incrementWidth * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                            }

                            if (incrementHeight != 0.0f)
                            {
                                dupShape.ScaleHeight((1.0f + (incrementHeight * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                            }

                            if (incrementRotation != 0.0f)
                            {
                                dupShape.Rotation += (incrementRotation * i);
                            }

                            if (incrementLeft != 0.0f)
                            {
                                dupShape.Left += (incrementLeft * i);
                            }

                            if (incrementTop != 0.0f)
                            {
                                dupShape.Top += (incrementTop * i);
                            }

                            if (incrementFont != 0.0f)
                            {
                                dupShape.TextFrame.TextRange.Font.Size += (incrementFont * i);
                            }

                            if (i == 1)
                            {
                                PowerPoint.Effect appear = sequence.AddEffect(dupShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                            }
                            else
                            {
                                PowerPoint.Effect appear = sequence.AddEffect(dupShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                appear.Timing.TriggerDelayTime = ((defaultDuration / numFrames) * i);
                            }

                            PowerPoint.Effect disappear = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            disappear.Exit = Office.MsoTriState.msoTrue;
                            disappear.Timing.TriggerDelayTime = ((defaultDuration / numFrames) * i);

                            lastShape = dupShape;
                        }
                        PowerPoint.Effect disappearLast = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        disappearLast.Exit = Office.MsoTriState.msoTrue;
                        disappearLast.Timing.TriggerDelayTime = defaultDuration;
                    }
                    else
                    {
                        //Motion Effect
                        if ((finalX != initialX) || (finalY != initialY))
                        {
                            effectMotion = sequence.AddEffect(sh1, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                            PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
                            effectMotion.Timing.Duration = defaultDuration;
                            trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;

                            //Create VML path for the motion path
                            //This path needs to be a curved path to allow the user to edit points
                            float point1X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                            float point1Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                            float point2X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                            float point2Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                            float point3X = (finalX - initialX) / presentation.PageSetup.SlideWidth;
                            float point3Y = (finalY - initialY) / presentation.PageSetup.SlideHeight;
                            motion.MotionEffect.Path = "M 0 0 C " + point1X + " " + point1Y + " " + point2X + " " + point2Y + " " + point3X + " " + point3Y + " E";
                            effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
                            effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;
                        }

                        //Resize Effect
                        if ((finalWidth != initialWidth) || (finalHeight != initialHeight))
                        {
                            sh1.LockAspectRatio = Office.MsoTriState.msoFalse;
                            effectResize = sequence.AddEffect(sh1, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                            PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
                            effectResize.Timing.Duration = defaultDuration;

                            resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
                            resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

                            trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                        }

                        //Rotation Effect
                        if (finalRotation != initialRotation)
                        {
                            effectRotate = sequence.AddEffect(sh1, PowerPoint.MsoAnimEffect.msoAnimEffectSpin, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                            PowerPoint.AnimationBehavior rotate = effectRotate.Behaviors[1];
                            effectRotate.Timing.Duration = defaultDuration;
                            effectRotate.EffectParameters.Amount = GetMinimumRotation(initialRotation, finalRotation);
                            trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                        }
                    }

                    PowerPoint.Effect shape2Appear = sequence.AddEffect(sh2, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

                    PowerPoint.Effect shape1Disappear = sequence.AddEffect(sh1, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    shape1Disappear.Exit = Office.MsoTriState.msoTrue;
                }
                Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
                AddAckSlide();
            }
            catch (Exception e)
            {
                LogException(e, "AddInSlideAnimationButtonClick");
                throw;
            }
        }
    }
}
