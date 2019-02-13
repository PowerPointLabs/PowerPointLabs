using System;
using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;

namespace PowerPointLabs.TooltipsLab
{
    internal static class AttachTriggerAnimation
    {
        public static void AddTriggerAnimation(PowerPointSlide currentSlide, Selection selection)
        {
            try
            {
                ShapeRange selectedShapes = selection.ShapeRange;

                if (selectedShapes.Count < 2)
                {
                    throw new Exception("Please select more than one shape.");
                }

                Shape triggerShape = selectedShapes[1];

                List<Shape> shapesToAnimate = GetShapesToAnimate(selectedShapes);

                AddTriggerAnimation(currentSlide, triggerShape, shapesToAnimate);
            }
            catch (Exception)
            {
                
            }
        }

        private static List<Shape> GetShapesToAnimate(ShapeRange selectedShapes)
        {
            List<Shape> animatedShapes = new List<Shape>();

            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                animatedShapes.Add(selectedShapes[i]);
            }

            return animatedShapes;
        }

        private static void AddTriggerAnimation(PowerPointSlide currentSlide, Shape triggerShape, List<Shape> shapesToAnimate)
        {
            TimeLine timeline = currentSlide.TimeLine;
            MsoAnimEffect appearEffect = MsoAnimEffect.msoAnimEffectFade;
            Sequence sequence = timeline.InteractiveSequences.Add();
            for (int i = 0; i < shapesToAnimate.Count; i++)
            {
                Shape animationShape = shapesToAnimate[i];
                MsoAnimTriggerType triggerType;
                if (i == 0)
                {
                    triggerType = MsoAnimTriggerType.msoAnimTriggerOnShapeClick;
                    sequence.AddTriggerEffect(animationShape, appearEffect, triggerType, triggerShape);
                }
                else
                {
                    triggerType = MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                    sequence.AddEffect(shapesToAnimate[i], appearEffect, MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                }
                if (i == 0)
                {
                    //triggerShape = shapesToAnimate[0];
                }
            }
        }

    }
}
