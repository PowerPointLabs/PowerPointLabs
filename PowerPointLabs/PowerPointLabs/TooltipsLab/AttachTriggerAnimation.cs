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
            MsoAnimEffect fadeEffect = MsoAnimEffect.msoAnimEffectFade;
            ISet<Shape> shapes = RemoveAnimationsInInteractiveSequence(currentSlide, triggerShape);
            Sequence sequence = timeline.InteractiveSequences.Add();
            // Append existing shapes to the list of shapes to animate
            shapesToAnimate.AddRange(shapes);
            // Add Entrance Effect to Shapes
            for (int i = 0; i < shapesToAnimate.Count; i++)
            {
                Shape animationShape = shapesToAnimate[i];
                MsoAnimTriggerType triggerType;
                // The first shape will be triggered by the click to appear
                if (i == 0)
                {
                    triggerType = MsoAnimTriggerType.msoAnimTriggerOnShapeClick;
                    sequence.AddTriggerEffect(animationShape, fadeEffect, triggerType, triggerShape);
                }
                // Rest of the shapes will appear with the first shape
                else
                {
                    triggerType = MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                    sequence.AddEffect(shapesToAnimate[i], fadeEffect, MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                }
            }

            // Add Exit Effect to Shapes
            for (int i = 0; i < shapesToAnimate.Count; i++)
            {
                Shape animationShape = shapesToAnimate[i];
                MsoAnimTriggerType triggerType;
                Effect effect;
                // The first shape will be triggered by the click to disappear
                if (i == 0)
                {
                    triggerType = MsoAnimTriggerType.msoAnimTriggerOnShapeClick;
                    effect = sequence.AddTriggerEffect(animationShape, fadeEffect, triggerType, triggerShape);
                }
                // Rest of the shapes will disappear with the first shape
                else
                {
                    triggerType = MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                    effect = sequence.AddEffect(shapesToAnimate[i], fadeEffect, MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                }
                effect.Exit = Microsoft.Office.Core.MsoTriState.msoTrue;
            }
        }

        private static ISet<Shape> RemoveAnimationsInInteractiveSequence(PowerPointSlide currentSlide, Shape triggerShape)
        {
            Sequences sequences = currentSlide.TimeLine.InteractiveSequences;
            // A set is used here so no duplicate shapes will be added
            ISet<Shape> existingAnimatedShapes = new HashSet<Shape>();
            // Find the existing sequence that has the triggerShape
            for (int i = 1; i <= sequences.Count; i++)
            {
                Sequence sequence = sequences[i];
                // Iterate from the back because we are deleting
                for (int j = sequence.Count; j >= 1; j--)
                {
                    Effect effect = sequence[i];
                    // Get existing shapes and only those with entrance effect
                    if (effect.Timing.TriggerShape == triggerShape)
                    {
                        existingAnimatedShapes.Add(effect.Shape);
                        effect.Delete();
                    }
                    else
                    {
                        break;
                    }
                }
            }

            return existingAnimatedShapes;
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
