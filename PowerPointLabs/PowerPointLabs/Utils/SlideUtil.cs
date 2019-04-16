using System.Collections.Generic;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;


namespace PowerPointLabs.Utils
{
    internal static class SlideUtil
    {
#pragma warning disable 0618

        #region API

        #region Slide Methods

        /// <summary>
        /// Sort by increasing index.
        /// </summary>
        public static void SortByIndex(List<PowerPointSlide> slides)
        {
            slides.Sort((sh1, sh2) => sh1.Index - sh2.Index);
        }

        /// <summary>
        /// Used for the SquashSlides method.
        /// This struct holds transition information for an effect.
        /// </summary>
        private struct EffectTransition
        {
            private readonly MsoAnimTriggerType slideTransition;
            private readonly float transitionTime;

            public EffectTransition(MsoAnimTriggerType slideTransition, float transitionTime)
            {
                this.slideTransition = slideTransition;
                this.transitionTime = transitionTime;
            }

            public void ApplyTransition(Effect effect)
            {
                effect.Timing.TriggerType = slideTransition;
                effect.Timing.TriggerDelayTime = transitionTime;
            }
        }

        /// <summary>
        /// Merges multiple animated slides into a single slide.
        /// TODO: Test this method more thoroughly, in places other than autozoom.
        /// </summary>
        public static void SquashSlides(IEnumerable<PowerPointSlide> slides)
        {
            PowerPointSlide firstSlide = null;
            ShapeRange previousShapes = null;
            EffectTransition slideTransition = new EffectTransition();

            foreach (PowerPointSlide slide in slides)
            {
                if (firstSlide == null)
                {
                    firstSlide = slide;
                    slideTransition = GetTransitionFromSlide(slide);

                    firstSlide.Transition.AdvanceOnClick = MsoTriState.msoTrue;
                    firstSlide.Transition.AdvanceOnTime = MsoTriState.msoFalse;

                    previousShapes = ShapeUtil.GetShapesWhenTypeNotMatches(firstSlide, firstSlide.Shapes.Range(), MsoShapeType.msoPlaceholder);
                    continue;
                }

                Sequence effectSequence = firstSlide.GetNativeSlide().TimeLine.MainSequence;
                int effectStartIndex = effectSequence.Count + 1;

                slide.DeleteIndicator();
                ShapeRange newShapeRange = firstSlide.CopyShapesToSlide(slide.Shapes.Range());
                newShapeRange.ZOrder(MsoZOrderCmd.msoSendToBack);

                foreach (Shape shape in newShapeRange)
                {
                    AddAppearAnimation(shape, firstSlide, effectStartIndex);
                }
                foreach (Shape shape in previousShapes)
                {
                    AddDisappearAnimation(shape, firstSlide, effectStartIndex);
                }
                slideTransition.ApplyTransition(effectSequence[effectStartIndex]);

                previousShapes = newShapeRange;
                slideTransition = GetTransitionFromSlide(slide);
                slide.Delete();
            }
        }

        #endregion

        #region Slide Design

        public static Design CreateDesign(string designName)
        {
            return PowerPointPresentation.Current.Presentation.Designs.Add(designName);
        }

        public static Design GetDesign(string designName)
        {
            foreach (Design design in PowerPointPresentation.Current.Presentation.Designs)
            {
                if (design.Name.Equals(designName))
                {
                    return design;
                }
            }
            return null;
        }

        public static void CopyToDesign(string designName, PowerPointSlide refSlide)
        {
            Design design = GetDesign(designName);
            if (design != null)
            {
                try
                {
                    design.Delete();
                } 
                catch (COMException e) 
                {
                    Logger.LogException(e, "CopyToDesign: Design cannot be deleted.");
                }
            }
            Design newDesign = PowerPointPresentation.Current.Presentation.Designs.Clone(refSlide.Design);
            newDesign.Name = designName;
        }

        # endregion

        #endregion

        #region Helper Methods

        /// <summary>
        /// Extracts the transition animation out of slide to be used as a transition animation for shapes.
        /// For now, it only extracts the trigger type (trigger by wait or by mouse click), not actual slide transitions.
        /// </summary>
        private static EffectTransition GetTransitionFromSlide(PowerPointSlide slide)
        {
            SlideShowTransition transition = slide.GetNativeSlide().SlideShowTransition;

            if (transition.AdvanceOnTime == MsoTriState.msoTrue)
            {
                return new EffectTransition(MsoAnimTriggerType.msoAnimTriggerAfterPrevious, transition.AdvanceTime);
            }
            return new EffectTransition(MsoAnimTriggerType.msoAnimTriggerOnPageClick, 0);
        }

        private static void AddAppearAnimation(Shape shape, PowerPointSlide inSlide, int effectStartIndex)
        {
            if (inSlide.HasEntryAnimation(shape))
            {
                return;
            }

            Effect effectFade = inSlide.GetNativeSlide().TimeLine.MainSequence.AddEffect(shape, MsoAnimEffect.msoAnimEffectAppear,
                MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious, effectStartIndex);
            effectFade.Exit = MsoTriState.msoFalse;
        }

        private static void AddDisappearAnimation(Shape shape, PowerPointSlide inSlide, int effectStartIndex)
        {
            if (inSlide.HasExitAnimation(shape))
            {
                return;
            }

            Effect effectFade = inSlide.GetNativeSlide().TimeLine.MainSequence.AddEffect(shape, MsoAnimEffect.msoAnimEffectAppear,
                MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious, effectStartIndex);
            effectFade.Exit = MsoTriState.msoTrue;
        }

        #endregion
    }
}
