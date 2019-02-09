using System;
using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.TooltipsLab
{
    internal static class AttachTriggerAnimation
    {
        public static void AddTriggerAnimation(PowerPointSlide currentSlide, Selection selection)
        {
            try
            {
                ShapeRange selectedShapes = selection.ShapeRange;

                AddTriggerAnimation(currentSlide, selectedShapes);
            }
            catch (Exception)
            {

            }
        }


        private static void AddTriggerAnimation(PowerPointSlide currentSlide, ShapeRange shapes)
        {
            MsoAnimTriggerType trigger = MsoAnimTriggerType.msoAnimTriggerOnPageClick;

            foreach (Shape animationShape in shapes)
            {
                Effect effectRotate = currentSlide.TimeLine.MainSequence.AddEffect(animationShape, MsoAnimEffect.msoAnimEffectSpin, MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                AnimationBehavior rotate = effectRotate.Behaviors[1];
                effectRotate.Timing.Duration = 1.0F;
                effectRotate.EffectParameters.Amount = LegacyShapeUtil.GetMinimumRotation(0.0F, 180.0F);
            }

        }

    }
}
