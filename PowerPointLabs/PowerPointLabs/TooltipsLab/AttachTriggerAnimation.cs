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
            if (shapes.Count < 2)
            {
                throw new Exception("Please use at least 2 shapes.");
            }

            Shape triggerShape = shapes[1];

            for (int i = 2; i <= shapes.Count; i++)
            {
                Shape animationShape = shapes[i];
                MsoAnimEffect appearEffect = MsoAnimEffect.msoAnimEffectFade;
                MsoAnimTriggerType triggerOnShapeClick = MsoAnimTriggerType.msoAnimTriggerOnShapeClick;
                TimeLine timeline = currentSlide.TimeLine;
                Sequence sequence = timeline.InteractiveSequences.Add();
                sequence.AddTriggerEffect(animationShape, appearEffect, triggerOnShapeClick, triggerShape);
            }

        }

    }
}
