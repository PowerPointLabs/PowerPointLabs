using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointSlide
    {
        protected readonly Slide _slide;
        private List<PowerPoint.MsoAnimEffect> entryEffects = new List<PowerPoint.MsoAnimEffect>()
        {
            PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimEffect.msoAnimEffectBlinds, PowerPoint.MsoAnimEffect.msoAnimEffectBox,
            PowerPoint.MsoAnimEffect.msoAnimEffectCheckerboard, PowerPoint.MsoAnimEffect.msoAnimEffectCircle, PowerPoint.MsoAnimEffect.msoAnimEffectDiamond,
            PowerPoint.MsoAnimEffect.msoAnimEffectDissolve, PowerPoint.MsoAnimEffect.msoAnimEffectFly, PowerPoint.MsoAnimEffect.msoAnimEffectPeek, 
            PowerPoint.MsoAnimEffect.msoAnimEffectPlus, PowerPoint.MsoAnimEffect.msoAnimEffectRandomBars, PowerPoint.MsoAnimEffect.msoAnimEffectSplit,
            PowerPoint.MsoAnimEffect.msoAnimEffectStrips, PowerPoint.MsoAnimEffect.msoAnimEffectWedge, PowerPoint.MsoAnimEffect.msoAnimEffectWheel,
            PowerPoint.MsoAnimEffect.msoAnimEffectWipe, PowerPoint.MsoAnimEffect.msoAnimEffectExpand, PowerPoint.MsoAnimEffect.msoAnimEffectFade,
            PowerPoint.MsoAnimEffect.msoAnimEffectFadedSwivel, PowerPoint.MsoAnimEffect.msoAnimEffectFadedZoom, PowerPoint.MsoAnimEffect.msoAnimEffectZoom,
            PowerPoint.MsoAnimEffect.msoAnimEffectCenterRevolve, PowerPoint.MsoAnimEffect.msoAnimEffectFloat, PowerPoint.MsoAnimEffect.msoAnimEffectGrowAndTurn,
            PowerPoint.MsoAnimEffect.msoAnimEffectRiseUp, PowerPoint.MsoAnimEffect.msoAnimEffectSpinner, PowerPoint.MsoAnimEffect.msoAnimEffectSwivel,
            PowerPoint.MsoAnimEffect.msoAnimEffectBoomerang, PowerPoint.MsoAnimEffect.msoAnimEffectBounce, PowerPoint.MsoAnimEffect.msoAnimEffectCredits,
            PowerPoint.MsoAnimEffect.msoAnimEffectFlip, PowerPoint.MsoAnimEffect.msoAnimEffectFloat, PowerPoint.MsoAnimEffect.msoAnimEffectPinwheel,
            PowerPoint.MsoAnimEffect.msoAnimEffectSpiral, PowerPoint.MsoAnimEffect.msoAnimEffectWhip
        };

        protected PowerPointSlide(Slide slide)
        {
            _slide = slide;
        }

        public static PowerPointSlide FromSlideFactory(Slide slide)
        {
            if (slide == null)
            {
                return null;
            }

            if (slide.Name.Contains("PPTLabsSpotlight"))
                return PowerPointSpotlightSlide.FromSlideFactory(slide);
            else if (slide.Name.Contains("PPTLabsAck"))
                return PowerPointAckSlide.FromSlideFactory(slide);
            else
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

        public TimeLine TimeLine
        {
            get { return _slide.TimeLine; }
        }

        public string Name
        {
            get { return _slide.Name; }
            set { _slide.Name = value; }
        }

        public void Delete()
        {
            _slide.Delete();
        }

        public void Copy()
        {
            _slide.Copy();
        }

        public void MoveTo(int index)
        {
            _slide.MoveTo(index);
        }

        public PowerPointSlide Duplicate()
        {
            Slide duplicatedSlide = _slide.Duplicate()[1];
            return PowerPointSlide.FromSlideFactory(duplicatedSlide);
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

        public void DeleteAllShapes()
        {
            List<Shape> shapes = _slide.Shapes.Cast<Shape>().ToList();

            var matchingShapes = shapes;
            foreach (Shape s in matchingShapes)
            {
                s.Delete();
            }
        }
        public void SetAudioAsAutoplay(Shape shape)
        {
            var mainSequence = _slide.TimeLine.MainSequence;

            Effect firstClickEvent = mainSequence.FindFirstAnimationForClick(1);
            bool hasNoClicksOnSlide = firstClickEvent == null;

            if (hasNoClicksOnSlide)
            {
                AddShapeAsLastAutoplaying(shape, MsoAnimEffect.msoAnimEffectMediaPlay);
            }
            else
            {
                InsertAnimationBeforeExisting(shape, firstClickEvent, MsoAnimEffect.msoAnimEffectMediaPlay);
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

        public Effect SetShapeAsClickTriggered(Shape shape, int clickNumber, MsoAnimEffect effect)
        {
            Effect addedEffect;

            Sequence mainSequence = _slide.TimeLine.MainSequence;
            Effect nextClickEffect = mainSequence.FindFirstAnimationForClick(clickNumber + 1);
            Effect previousClickEffect = mainSequence.FindFirstAnimationForClick(clickNumber);

            bool hasClicksAfter = nextClickEffect != null;
            bool hasClickBefore = previousClickEffect != null;

            if (hasClicksAfter)
            {
                addedEffect = InsertAnimationBeforeExisting(shape, nextClickEffect, effect);
            }
            else if (hasClickBefore)
            {
                addedEffect = AddShapeAsLastAutoplaying(shape, effect);
            }
            else
            {
                addedEffect = mainSequence.AddEffect(shape, effect);
            }

            return addedEffect;
        }

        private Effect AddShapeAsLastAutoplaying(Shape shape, MsoAnimEffect effect)
        {
            Effect addedEffect = _slide.TimeLine.MainSequence.AddEffect(shape, effect,
                MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            return addedEffect;
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

        public void DeleteShapeAnimations(Shape sh)
        {
            PowerPoint.Sequence sequence = _slide.TimeLine.MainSequence;
            for (int x = sequence.Count; x >= 1; x--)
            {
                PowerPoint.Effect effect = sequence[x];
                if (effect.Shape.Name == sh.Name && effect.Shape.Id == sh.Id)
                    effect.Delete();
            }
        }

        public void RemoveAnimationsForShapes(List<Shape> shapes)
        {
            foreach (PowerPoint.Shape sh in shapes)
            {
                DeleteShapeAnimations(sh);
            }
        }

        public List<PowerPoint.Shape> GetShapesWithPrefix(string prefix)
        {
            List<Shape> shapes = _slide.Shapes.Cast<Shape>().ToList();
            List<Shape> matchingShapes = shapes.Where(current => current.Name.StartsWith(prefix)).ToList();

            return matchingShapes;
        }

        private static void DeleteEffectsForShape(Shape shape, IEnumerable<Effect> mainEffects)
        {
            foreach (Effect e in mainEffects.Where(e => e.Shape.Equals(shape)))
            {
                e.Delete();
            }
        }

        public PowerPointSlide CreateSpotlightSlide()
        {
            Slide duplicatedSlide = _slide.Duplicate()[1];
            return PowerPointSpotlightSlide.FromSlideFactory(duplicatedSlide);
        }

        public PowerPointSlide CreateAutoAnimateSlide()
        {
            Slide duplicatedSlide = _slide.Duplicate()[1];
            return PowerPointAutoAnimateSlide.FromSlideFactory(duplicatedSlide);
        }

        public PowerPointSlide CreateDrillDownSlide()
        {
            //Slide addedSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(_slide.SlideIndex + 1, _slide.CustomLayout);
            Slide duplicatedSlide = _slide.Duplicate()[1];
            return PowerPointDrillDownSlide.FromSlideFactory(duplicatedSlide);
        }

        public PowerPointSlide CreateStepBackSlide()
        {
            Slide duplicatedSlide = _slide.Duplicate()[1];
            return PowerPointStepBackSlide.FromSlideFactory(duplicatedSlide);
        }

        public PowerPointSlide CreateZoomToAreaSingleSlide()
        {
            Slide duplicatedSlide = _slide.Duplicate()[1];
            return PowerPointZoomToAreaSingleSlide.FromSlideFactory(duplicatedSlide);
        }

        public PowerPointSlide CreateZoomMagnifyingSlide()
        {
            Slide duplicatedSlide = _slide.Duplicate()[1];
            return PowerPointMagnifyingSlide.FromSlideFactory(duplicatedSlide);
        }

        public PowerPointSlide CreateZoomMagnifiedSlide()
        {
            Slide duplicatedSlide = _slide.Duplicate()[1];
            return PowerPointMagnifiedSlide.FromSlideFactory(duplicatedSlide);
        }

        public PowerPointSlide CreateZoomDeMagnifyingSlide()
        {
            Slide duplicatedSlide = _slide.Duplicate()[1];
            return PowerPointDeMagnifyingSlide.FromSlideFactory(duplicatedSlide);
        }

        public PowerPointSlide CreateZoomPanSlide()
        {
            Slide duplicatedSlide = _slide.Duplicate()[1];
            return PowerPointMagnifiedPanSlide.FromSlideFactory(duplicatedSlide);
        }
        public bool HasExitAnimation(Shape shape)
        {
            PowerPoint.Sequence sequence = _slide.TimeLine.MainSequence;
            for (int x = sequence.Count; x >= 1; x--)
            {
                PowerPoint.Effect effect = sequence[x];
                if (effect.Shape.Name == shape.Name && effect.Shape.Id == shape.Id)
                    if (effect.Exit == Office.MsoTriState.msoTrue)
                        return true;
            }
            return false;
        }

        public bool HasEntryAnimation(PowerPoint.Shape shape)
        {
            PowerPoint.Sequence sequence = _slide.TimeLine.MainSequence;
            for (int x = sequence.Count; x >= 1; x--)
            {
                PowerPoint.Effect effect = sequence[x];
                if (effect.Shape.Name == shape.Name && effect.Shape.Id == shape.Id)
                    if (entryEffects.Contains(effect.EffectType))
                        return true;
            }
            return false;
        }

        protected void DeleteSlideNotes()
        {
            if (_slide.HasNotesPage == MsoTriState.msoTrue)
            {
                foreach (Shape sh in _slide.NotesPage.Shapes)
                {
                    if (sh.TextFrame.HasText == MsoTriState.msoTrue)
                        sh.TextEffect.Text = "";
                }
            }
        }

        protected void DeleteSlideMedia()
        {
            foreach (Shape sh in _slide.Shapes)
            {
                if (sh.Type == MsoShapeType.msoMedia)
                    sh.Delete();
            }
        }

        protected Shape AddPowerPointLabsIndicator()
        {
            String tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            Shape indicatorShape = _slide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, PowerPointPresentation.SlideWidth - 120, 0, 120, 84);
            
            indicatorShape.Left = PowerPointPresentation.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.Name = "PPTLabsIndicator" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

            PowerPoint.Effect effectAppear = null;
            PowerPoint.Effect effectDisappear = null;
            PowerPoint.Sequence sequence = _slide.TimeLine.MainSequence;

            effectAppear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectAppear.Timing.Duration = 0;

            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;

            return indicatorShape;
        }

        protected void RemoveSlideTransitions()
        {
            _slide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
        }

        public void MoveMotionAnimation()
        {
            PowerPoint.Sequence sequence = _slide.TimeLine.MainSequence;
            foreach (PowerPoint.Effect eff in _slide.TimeLine.MainSequence)
            {
                if ((eff.EffectType >= PowerPoint.MsoAnimEffect.msoAnimEffectPathCircle && eff.EffectType <= PowerPoint.MsoAnimEffect.msoAnimEffectPathRight) || eff.EffectType == PowerPoint.MsoAnimEffect.msoAnimEffectCustom)
                {
                    PowerPoint.AnimationBehavior motion = eff.Behaviors[1];
                    if (motion.Type == PowerPoint.MsoAnimType.msoAnimTypeMotion)
                    {
                        PowerPoint.Shape sh = eff.Shape;
                        string motionPath = motion.MotionEffect.Path.Trim();
                        if (motionPath.Last() < 'A' || motionPath.Last() > 'Z')
                            motionPath += " X";
                        string[] path = motionPath.Split(' ');
                        int count = path.Length;
                        float xVal = Convert.ToSingle(path[count - 3]);
                        float yVal = Convert.ToSingle(path[count - 2]);
                        sh.Left += (xVal * PowerPointPresentation.SlideWidth);
                        sh.Top += (yVal * PowerPointPresentation.SlideHeight);
                    }
                }
            }
        }

        public void AddAppearDisappearAnimation(Shape sh)
        {
            PowerPoint.Sequence sequence = _slide.TimeLine.MainSequence;
            PowerPoint.Effect effectAppear = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectAppear.Timing.Duration = 0;

            PowerPoint.Effect effectDisappear = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;
        }

        public Shape GetShapeWithSameIDAndName(Shape shapeToMatch)
        {
            Shape tempMatchingShape = null;
            foreach (Shape sh in _slide.Shapes)
            {
                if (shapeToMatch.Id == sh.Id && haveSameNames(shapeToMatch, sh))
                {
                    if (tempMatchingShape == null)
                        tempMatchingShape = sh;
                    else
                    {
                        if (GetDistanceBetweenShapes(shapeToMatch, sh) < GetDistanceBetweenShapes(shapeToMatch, tempMatchingShape))
                            tempMatchingShape = sh;
                    }
                }
            }
            return tempMatchingShape;
        }

        public Shape GetShapeWithSameName(Shape shapeToMatch)
        {
            Shape tempMatchingShape = null;
            foreach (Shape sh in _slide.Shapes)
            {
                if (haveSameNames(shapeToMatch, sh))
                {
                    if (tempMatchingShape == null)
                        tempMatchingShape = sh;
                    else
                    {
                        if (GetDistanceBetweenShapes(shapeToMatch, sh) < GetDistanceBetweenShapes(shapeToMatch, tempMatchingShape))
                            tempMatchingShape = sh;
                    }
                }
            }
            return tempMatchingShape;
        }

        private float GetDistanceBetweenShapes(PowerPoint.Shape sh1, PowerPoint.Shape sh2)
        {
            float sh1CenterX = (sh1.Left + (sh1.Width / 2));
            float sh2CenterX = (sh2.Left + (sh2.Width / 2));
            float sh1CenterY = (sh1.Top + (sh1.Height / 2));
            float sh2CenterY = (sh2.Top + (sh2.Height / 2));
            float distSquared = (float)(Math.Pow((sh2CenterX - sh1CenterX), 2) + Math.Pow((sh2CenterY - sh1CenterY), 2));
            return (float)(Math.Sqrt(distSquared));
        }

        private bool haveSameNames(PowerPoint.Shape sh1, PowerPoint.Shape sh2)
        {
            String name1 = sh1.Name;
            String name2 = sh2.Name;

            return (name1.ToUpper().CompareTo(name2.ToUpper()) == 0);
        }

        public bool isSpotlightSlide()
        {
            return _slide.Name.Contains("PPTLabsSpotlight");
        }

        public bool isAckSlide()
        {
            return _slide.Name.Contains("PPTLabsAck");
        }

        public PowerPointSlide CreateAckSlide()
        {
            Slide ackSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides.Add(Globals.ThisAddIn.Application.ActivePresentation.Slides.Count + 1, PpSlideLayout.ppLayoutBlank);
            return PowerPointAckSlide.FromSlideFactory(ackSlide);
        }
    }
}
