using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace PowerPointLabs.Models
{
    class PowerPointSlide
    {
        private readonly Slide _slide;

        private PowerPointSlide(Slide slide)
        {
            _slide = slide;
        }

        public static PowerPointSlide FromSlideFactory(Slide slide)
        {
            if (slide == null)
            {
                return null;
            }

            return new PowerPointSlide(slide);
        }

        public String NotesPageText
        {
            get
            {
                if (_slide == null || _slide.HasNotesPage == MsoTriState.msoFalse)
                {
                    return String.Empty;
                }

                IEnumerable<Shape> notesPagePlaceholders = _slide.NotesPage.Shapes.Placeholders.Cast<Shape>();
                Shape notesPageBody = notesPagePlaceholders.FirstOrDefault(shape => shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody);

                String notesText = notesPageBody != null ? notesPageBody.TextFrame.TextRange.Text : String.Empty;
                return notesText;
            }
        }

        public Shapes Shapes
        {
            get { return _slide.Shapes; }
        }

        public int Index
        {
            get { return _slide.SlideIndex; }
        }

        public SlideShowTransition Transition
        {
            get { return _slide.SlideShowTransition; }
        }

        public void DeleteShapesWithPrefix(string prefix)
        {
            List<Shape> shapes = _slide.Shapes.Cast<Shape>().ToList();

            var matchingShapes = shapes.Where(current => current.Name.StartsWith(prefix));
            foreach (Shape s in matchingShapes)
            {
                s.Delete();
            }
        }

        public void SetAudioAsAutoplay(Shape shape)
        {
            var shapesTriggeredByClick = GetShapesTriggeredByClick();

            if (shapesTriggeredByClick.Count == 0)
            {
                AddShapeAsLastAutoplaying(shape, MsoAnimEffect.msoAnimEffectMediaPlay);
            }
            else
            {
                var shapeClickEffect = _slide.TimeLine.MainSequence.FindFirstAnimationForClick(1);

                InsertAnimationBeforeExisting(shape, shapeClickEffect, MsoAnimEffect.msoAnimEffectMediaPlay);
            }
        }

        private Effect InsertAnimationBeforeExisting(Shape shape, Effect existing, MsoAnimEffect effect)
        {
            var sequence = _slide.TimeLine.MainSequence;

            Effect newAnimation = sequence.AddEffect(shape, effect, MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            newAnimation.MoveBefore(existing);

            return newAnimation;
        }

        public Effect SetShapeAsClickTriggered(Shape shape, int index, MsoAnimEffect effect)
        {
            var shapesTriggeredByClick = GetShapesTriggeredByClick();

            Effect addedEffect;
            int clickIndex = index + 1; // Clicks are 1-indexed while shapes are 0-indexed.
            if (shapesTriggeredByClick.Count > clickIndex) 
            {
                var clickEffect = _slide.TimeLine.MainSequence.FindFirstAnimationForClick(clickIndex);

                addedEffect = InsertAnimationBeforeExisting(shape, clickEffect, effect);
            }
            else if (shapesTriggeredByClick.Count == index)
            {
                addedEffect = AddShapeAsLastAutoplaying(shape, effect);
            }
            else
            {
                // Just add this as a new click effect.
                var animationSequence = _slide.TimeLine.MainSequence;
                addedEffect = animationSequence.AddEffect(shape, effect);
            }

            return addedEffect;
        }

        private Effect AddShapeAsLastAutoplaying(Shape shape, MsoAnimEffect effect)
        {
            Effect addedEffect = _slide.TimeLine.MainSequence.AddEffect(shape, effect,
                MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            return addedEffect;
        }

        private Effect InsertAnimationAtIndex(Shape shape, int index, MsoAnimEffect animationEffect)
        {
            var animationSequence = _slide.TimeLine.MainSequence;
            Effect effect = animationSequence.AddEffect(shape, animationEffect, MsoAnimateByLevel.msoAnimateLevelNone,
                MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effect.MoveTo(index);
            return effect;
        }

        private Effect InsertAnimationAtIndex(Shape shape, int index, MsoAnimEffect animationEffect,
            MsoAnimTriggerType triggerType)
        {
            var animationSequence = _slide.TimeLine.MainSequence;
            Effect effect = animationSequence.AddEffect(shape, animationEffect, MsoAnimateByLevel.msoAnimateLevelNone,
                triggerType);
            effect.MoveTo(index);
            return effect;
        }

        private List<Shape> GetShapesTriggeredByClick()
        {
            var shapesWithAnimations = GetShapesWithAnimations();
            var shapesOnClick =
                shapesWithAnimations.Where(shape => shape.AnimationSettings.AdvanceMode == PpAdvanceMode.ppAdvanceOnClick)
                    .ToList();
            return shapesOnClick;
        }

        private IEnumerable<Shape> GetShapesWithAnimations()
        {
            var shapesWithAnimations = _slide.TimeLine.MainSequence.Cast<Effect>().Select(effect => effect.Shape).ToList();
            return shapesWithAnimations;
        }

        public void ShowShapeAfterClick(Shape shape, int clickNumber)
        {
            SetShapeAsClickTriggered(shape, clickNumber, MsoAnimEffect.msoAnimEffectAppear);
        }

        public void HideShapeAfterClick(Shape shape, int clickNumber)
        {
            Effect addedEffect = SetShapeAsClickTriggered(shape, clickNumber, MsoAnimEffect.msoAnimEffectAppear);
            addedEffect.Exit = MsoTriState.msoTrue;
        }

        public void HideShapeAsLastClickIfNeeded(Shape shape)
        {
            if (IsNextSlideTransitionBlacklisted())
            {
                var animationSequence = _slide.TimeLine.MainSequence;
                var effect = animationSequence.AddEffect(shape, MsoAnimEffect.msoAnimEffectFade);
                effect.Exit = MsoTriState.msoTrue;
            }
        }

        private bool IsNextSlideTransitionBlacklisted()
        {
            bool isLastSlide = _slide.SlideIndex == PowerPointPresentation.SlideCount;
            if (isLastSlide)
            {
                return false;
            }

            // Indexes are from 1, while the slide collection starts from 0.
            PowerPointSlide nextSlide = PowerPointPresentation.Slides.ElementAt(Index);
            switch (nextSlide.Transition.EntryEffect)
            {
                case PpEntryEffect.ppEffectCoverUp:
                case PpEntryEffect.ppEffectCoverLeftUp:
                case PpEntryEffect.ppEffectCoverRightUp:
                case PpEntryEffect.ppEffectFlyFromBottom:
                case PpEntryEffect.ppEffectPushUp:
                case PpEntryEffect.ppEffectPushDown:
                case PpEntryEffect.ppEffectSwitchUp:
                case PpEntryEffect.ppEffectFlipUp:
                case PpEntryEffect.ppEffectCubeUp:
                case PpEntryEffect.ppEffectRotateUp:
                case PpEntryEffect.ppEffectBoxUp:
                case PpEntryEffect.ppEffectOrbitUp:
                case PpEntryEffect.ppEffectPanUp:
                    return true;
                default:
                    return false;
            }
        }

        public void TransferAnimation(Shape source, Shape destination)
        {
            Sequence sequence = _slide.TimeLine.MainSequence;
            var enumerableSequence = sequence.Cast<Effect>().ToList();

            Effect entryDetails = enumerableSequence.FirstOrDefault(effect => effect.Shape.Equals(source));
            if (entryDetails != null)
            {
                InsertAnimationAtIndex(destination, entryDetails.Index, entryDetails.EffectType, entryDetails.Timing.TriggerType);
            }

            Effect exitDetails = enumerableSequence.Last(effect => effect.Shape.Equals(source));
            if (exitDetails != null && !exitDetails.Equals(entryDetails))
            {
                InsertAnimationAtIndex(destination, exitDetails.Index, exitDetails.EffectType,
                    exitDetails.Timing.TriggerType);
            }
        }

        public void RemoveAnimationsForShape(Shape shape)
        {
            IEnumerable<Effect> mainEffects = _slide.TimeLine.MainSequence.Cast<Effect>();
            DeleteEffectsForShape(shape, mainEffects);

            foreach (Sequence sequence in _slide.TimeLine.InteractiveSequences)
            {
                DeleteEffectsForShape(shape, sequence.Cast<Effect>());
            }
        }

        private static void DeleteEffectsForShape(Shape shape, IEnumerable<Effect> mainEffects)
        {
            foreach (Effect e in mainEffects.Where(e => e.Shape.Equals(shape)))
            {
                e.Delete();
            }
        }
    }
}
