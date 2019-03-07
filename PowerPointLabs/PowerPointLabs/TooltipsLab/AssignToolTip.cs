using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.TooltipsLab
{
    /// <summary>
    /// Assigns tooltip to shapes, follows the convention that the first shape is the trigger shape, the rest
    /// of the shapes will be the callouts. All callout shapes will appear on first click of the trigger shape,
    /// and all will disappear on second click of the trigger shape.
    /// </summary>
    internal static class AssignTooltip
    {
        public static bool AddTriggerAnimation(PowerPointSlide currentSlide, Selection selection)
        {
            ShapeRange selectedShapes = selection.ShapeRange;
            
            if (selectedShapes.Count < 2)
            {
                MessageBox.Show(TooltipsLabText.ErrorLessThanTwoShapesSelected,
                    TooltipsLabText.ErrorTooltipsDialogTitle);

                return false;
            }

            Shape triggerShape = selectedShapes[1];

            List<Shape> shapesToAnimate = GetShapesToAnimate(selectedShapes);

            AddTriggerAnimation(currentSlide, triggerShape, shapesToAnimate);

            return true;
        }

        public static void AddTriggerAnimation(PowerPointSlide currentSlide, Shape triggerShape, Shape calloutShape)
        {
            List<Shape> calloutShapeList = new List<Shape>();
            calloutShapeList.Add(calloutShape);
            AddTriggerAnimation(currentSlide, triggerShape, calloutShapeList);
        }

        private static void AddTriggerAnimation(PowerPointSlide currentSlide, Shape triggerShape, List<Shape> newShapesToAnimate)
        {
            TimeLine timeline = currentSlide.TimeLine;

            // Get the shapes that are already associated with trigger shape
            List<Shape> shapesToAnimate = GetShapesInInteractiveSequenceWithAnimationsRemoved(currentSlide, triggerShape, newShapesToAnimate);
            Sequence sequence = timeline.InteractiveSequences.Add();

            AddTriggerEffect(triggerShape, shapesToAnimate, TooltipsLabSettings.AnimationType, sequence);
        }

        private static void AddTriggerEffect(Shape triggerShape, List<Shape> shapesToAnimate, MsoAnimEffect effect, Sequence sequence)
        {
            // Add Entrance Effect
            for (int i = 0; i < shapesToAnimate.Count; i++)
            {
                Shape animationShape = shapesToAnimate[i];
                MsoAnimTriggerType triggerType;
                // The first shape will be triggered by the click to appear
                if (i == 0)
                {
                    triggerType = MsoAnimTriggerType.msoAnimTriggerOnShapeClick;
                    sequence.AddTriggerEffect(animationShape, effect, triggerType, triggerShape);
                }
                // Rest of the shapes will appear with the first shape
                else
                {
                    triggerType = MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                    sequence.AddEffect(shapesToAnimate[i], effect,
                        MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                }
            }

            // Add Exit Effect to Shapes
            for (int i = 0; i < shapesToAnimate.Count; i++)
            {
                Shape animationShape = shapesToAnimate[i];
                MsoAnimTriggerType triggerType;
                Effect effectInSequence;
                // The first shape will be triggered by the click to disappear
                if (i == 0)
                {
                    triggerType = MsoAnimTriggerType.msoAnimTriggerOnShapeClick;
                    effectInSequence = sequence.AddTriggerEffect(animationShape, effect, triggerType, triggerShape);
                }
                // Rest of the shapes will disappear with the first shape
                else
                {
                    triggerType = MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                    effectInSequence = sequence.AddEffect(shapesToAnimate[i], effect, 
                        MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                }
                effectInSequence.Exit = Microsoft.Office.Core.MsoTriState.msoTrue;
            }
        }

        private static List<Shape> GetShapesInInteractiveSequenceWithAnimationsRemoved(PowerPointSlide currentSlide, Shape triggerShape, List<Shape> shapesToAnimate)
        {
            Sequences sequences = currentSlide.TimeLine.InteractiveSequences;
            // A set is used here so no duplicate shapes will be added
            ISet<Shape> shapesToAnimateSet = new HashSet<Shape>(shapesToAnimate);

            // Find the existing sequence that has the triggerShape
            for (int i = 1; i <= sequences.Count; i++)
            {
                Sequence sequence = sequences[i];
                // Iterate from the back because of deletion
                for (int j = sequence.Count; j >= 1; j--)
                {
                    Effect effect = sequence[j];
                    // A sequence is attached to a trigger shape. However we can only use the effect to find out
                    // what is the trigger shape, thus we break when the first effect's trigger shape is not 
                    // what we are looking for and delete all effects from the sequence otherwise.
                    if (effect.Timing.TriggerShape == triggerShape)
                    {
                        shapesToAnimateSet.Add(effect.Shape);
                        effect.Delete();
                    }
                    else
                    {
                        break;
                    }
                }
            }

            return new List<Shape>(shapesToAnimateSet);
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
