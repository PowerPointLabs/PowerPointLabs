using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.TooltipsLab
{
    internal static class AssignTooltip
    {
        public static void AddTriggerAnimation(PowerPointSlide currentSlide, Selection selection)
        {
            ShapeRange selectedShapes = selection.ShapeRange;

            if (selectedShapes.Count < 2)
            {
                MessageBox.Show(TooltipsLabText.ErrorLessThanTwoShapesSelected,
                    TooltipsLabText.ErrorTooltipsDialogTitle);

                // TODO: New Exception for TooltipsLab
                throw new Exception();
            }

            try
            {
                Shape triggerShape = selectedShapes[1];

                List<Shape> shapesToAnimate = GetShapesToAnimate(selectedShapes);

                AddTriggerAnimation(currentSlide, triggerShape, shapesToAnimate);
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        public static void AddTriggerAnimation(PowerPointSlide currentSlide, Shape triggerShape, Shape calloutShape)
        {
            List<Shape> calloutShapeList = new List<Shape>();
            calloutShapeList.Add(calloutShape);
            AddTriggerAnimation(currentSlide, triggerShape, calloutShapeList);
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
                // The first shape will be triggered by the click
                if (i == 0)
                {
                    triggerType = MsoAnimTriggerType.msoAnimTriggerOnShapeClick;
                    sequence.AddTriggerEffect(animationShape, appearEffect, triggerType, triggerShape);
                }
                // Rest of the shapes will appear with the first shape
                else
                {
                    triggerType = MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                    sequence.AddEffect(shapesToAnimate[i], appearEffect, MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                }
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
    }
}
