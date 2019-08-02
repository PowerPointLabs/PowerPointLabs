using System;
using System.Collections.Generic;
using System.Linq;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.AnimationLab;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointMagnifiedPanSlide : PowerPointSlide
    {
        private PowerPoint.Shape indicatorShape = null;
        private PowerPoint.Shape panShapeFrom = null;
        private PowerPoint.Shape panShapeTo = null;

        private PowerPointMagnifiedPanSlide(PowerPoint.Slide slide) : base(slide)
        {
            _slide.Name = "PPTLabsMagnifiedPanSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
        }

        public static PowerPointSlide FromSlideFactory(PowerPoint.Slide slide)
        {
            if (slide == null)
            {
                return null;
            }

            return new PowerPointMagnifiedPanSlide(slide);
        }

        public void AddZoomToAreaAnimation(PowerPointSlide slideToPanFrom, PowerPointSlide slideToPanTo)
        {
            PrepareForZoomToArea(slideToPanFrom, slideToPanTo);
            DefaultMotionAnimation.AddZoomToAreaPanAnimation(this, panShapeFrom, panShapeTo, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
            DefaultMotionAnimation.PreloadShape(this, panShapeFrom);
            
            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
        }

        private void PrepareForZoomToArea(PowerPointSlide slideToPanFrom, PowerPointSlide slideToPanTo)
        {
            //Delete all shapes from slide excpet last magnified shape
            List<PowerPoint.Shape> shapes = _slide.Shapes.Cast<PowerPoint.Shape>().ToList();
            IEnumerable<PowerPoint.Shape> matchingShapes = shapes.Where(current => (!current.Name.Contains("PPTLabsMagnifyAreaGroup")));
            foreach (PowerPoint.Shape s in matchingShapes)
            {
                s.SafeDelete();
            }

            panShapeFrom = GetShapesWithPrefix("PPTLabsMagnifyAreaGroup")[0];
            panShapeTo = slideToPanTo.GetShapesWithPrefix("PPTLabsMagnifyAreaGroup")[0];

            //Add fade animation to existing shapes
            shapes = _slide.Shapes.Cast<PowerPoint.Shape>().ToList();
            matchingShapes = shapes.Where(current => (!(current.Equals(indicatorShape) || current.Equals(panShapeFrom))));
            foreach (PowerPoint.Shape s in matchingShapes)
            {
                DeleteShapeAnimations(s);
                PowerPoint.Effect effectFade = _slide.TimeLine.MainSequence.AddEffect(s, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                effectFade.Exit = Office.MsoTriState.msoTrue;
                effectFade.Timing.Duration = 0.25f;
            }

            DeleteSlideNotes();
            DeleteSlideMedia();
            ManageSlideTransitions();
            indicatorShape = AddPowerPointLabsIndicator();
        }

        private void ManageSlideTransitions()
        {
            RemoveSlideTransitions();
            _slide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
            _slide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
            _slide.SlideShowTransition.AdvanceTime = 0;
        }
    }
}
