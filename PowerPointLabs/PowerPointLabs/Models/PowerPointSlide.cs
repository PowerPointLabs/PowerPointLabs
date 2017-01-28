using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using Office = Microsoft.Office.Core;

namespace PowerPointLabs.Models
{
    public class PowerPointSlide
    {
#pragma warning disable 0618
        public const string PptLabsIndicatorShapeName = "PPTIndicator";
        private const string PptLabsTemplateMarkerShapeName = "PPTTemplateMarker";
        private const string UnnamedShapeName = "Unnamed Shape ";

        protected readonly Slide _slide;

        private List<MsoAnimEffect> entryEffects = new List<MsoAnimEffect>()
        {
            MsoAnimEffect.msoAnimEffectAppear, MsoAnimEffect.msoAnimEffectBlinds, MsoAnimEffect.msoAnimEffectBox,
            MsoAnimEffect.msoAnimEffectCheckerboard, MsoAnimEffect.msoAnimEffectCircle, MsoAnimEffect.msoAnimEffectDiamond,
            MsoAnimEffect.msoAnimEffectDissolve, MsoAnimEffect.msoAnimEffectFly, MsoAnimEffect.msoAnimEffectPeek, 
            MsoAnimEffect.msoAnimEffectPlus, MsoAnimEffect.msoAnimEffectRandomBars, MsoAnimEffect.msoAnimEffectSplit,
            MsoAnimEffect.msoAnimEffectStrips, MsoAnimEffect.msoAnimEffectWedge, MsoAnimEffect.msoAnimEffectWheel,
            MsoAnimEffect.msoAnimEffectWipe, MsoAnimEffect.msoAnimEffectExpand, MsoAnimEffect.msoAnimEffectFade,
            MsoAnimEffect.msoAnimEffectFadedSwivel, MsoAnimEffect.msoAnimEffectFadedZoom, MsoAnimEffect.msoAnimEffectZoom,
            MsoAnimEffect.msoAnimEffectCenterRevolve, MsoAnimEffect.msoAnimEffectFloat, MsoAnimEffect.msoAnimEffectGrowAndTurn,
            MsoAnimEffect.msoAnimEffectRiseUp, MsoAnimEffect.msoAnimEffectSpinner, MsoAnimEffect.msoAnimEffectSwivel,
            MsoAnimEffect.msoAnimEffectBoomerang, MsoAnimEffect.msoAnimEffectBounce, MsoAnimEffect.msoAnimEffectCredits,
            MsoAnimEffect.msoAnimEffectFlip, MsoAnimEffect.msoAnimEffectFloat, MsoAnimEffect.msoAnimEffectPinwheel,
            MsoAnimEffect.msoAnimEffectSpiral, MsoAnimEffect.msoAnimEffectWhip
        };

        protected PowerPointSlide(Slide slide)
        {
            _slide = slide;
        }

        public Slide GetNativeSlide()
        {
            return _slide;
        }

        public static PowerPointSlide FromSlideFactory(Slide slide, bool includeIndicator = false)
        {
            if (slide == null)
            {
                return null;
            }

            PowerPointSlide powerPointSlide;
            if (slide.Name.Contains("PPTLabsSpotlight"))
                powerPointSlide = PowerPointSpotlightSlide.FromSlideFactory(slide);
            else if (PowerPointAckSlide.IsAckSlide(slide))
                powerPointSlide = PowerPointAckSlide.FromSlideFactory(slide);
            else
                powerPointSlide = new PowerPointSlide(slide);

            if (includeIndicator)
            {
                powerPointSlide.AddPowerPointLabsIndicator();
            }
            return powerPointSlide;
        }

        public String NotesPageText
        {
            get
            {
                if (_slide == null || _slide.HasNotesPage == MsoTriState.msoFalse)
                {
                    return String.Empty;
                }

                var notesPagePlaceholders = _slide.NotesPage.Shapes.Placeholders.Cast<Shape>();
                var notesPageBody = notesPagePlaceholders.FirstOrDefault(shape => shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody);

                String notesText = notesPageBody != null ? notesPageBody.TextFrame.TextRange.Text : String.Empty;
                return notesText;
            }

            set
            {
                if (_slide == null)
                {
                    return;
                }

                var notesPagePlaceholders = _slide.NotesPage.Shapes.Placeholders.Cast<Shape>();
                var notesPageBody = notesPagePlaceholders.FirstOrDefault(shape => shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody);

                if (notesPageBody != null)
                {
                    notesPageBody.TextFrame.TextRange.Text = value;
                }
            }
        }

        /// <summary>
        /// TODO: escape newlines so that they can be stored properly?
        /// TODO: It is a known problem that if you store a string with newlines in NotesPageText, the retrieved string may be slightly different.
        /// </summary>
        public void StoreDataInNotes(string data)
        {
            NotesPageText = TextCollection.NotesPageStorageText + data;
        }

        public string RetrieveDataFromNotes()
        {
            var text = NotesPageText;
            if (!text.StartsWith(TextCollection.NotesPageStorageText)) return "";
            return text.Substring(TextCollection.NotesPageStorageText.Length);
        }

        public Shapes Shapes
        {
            get { return _slide.Shapes; }
        }

        public int ID
        {
            get { return _slide.SlideID; }
        }

        public int Index
        {
            get { return _slide.SlideIndex; }
        }

        public Design Design
        {
            get { return _slide.Design; }
            set { _slide.Design = value; }
        }

        public PpSlideLayout Layout
        {
            get { return _slide.Layout; }
            set { _slide.Layout = value; }
        }

        /// <summary>
        /// It only copies the background colour for now. Is there really no way to copy over background in general?
        /// </summary>
        public void CopyBackgroundColourFrom(PowerPointSlide refSlide)
        {
            var myFill = _slide.Background[1].Fill;
            var refFill = refSlide._slide.Background[1].Fill;

            if (refFill.Type == MsoFillType.msoFillSolid)
            {
                _slide.FollowMasterBackground = MsoTriState.msoFalse;
                myFill.ForeColor.RGB = refFill.ForeColor.RGB;
                myFill.BackColor.RGB = refFill.BackColor.RGB;
            }
        }

        public SlideShowTransition Transition
        {
            get { return _slide.SlideShowTransition; }
            set
            {
                // deep copy set-able fields
                _slide.SlideShowTransition.AdvanceOnClick = value.AdvanceOnClick;
                _slide.SlideShowTransition.AdvanceOnTime = value.AdvanceOnTime;
                _slide.SlideShowTransition.AdvanceTime = value.AdvanceTime;
                _slide.SlideShowTransition.Duration = value.Duration;
                _slide.SlideShowTransition.EntryEffect = value.EntryEffect;
                _slide.SlideShowTransition.Hidden = value.Hidden;
                _slide.SlideShowTransition.Speed = value.Speed;
            }
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

        public bool Hidden
        {
            get { return Transition.Hidden == MsoTriState.msoTrue; }
            set { Transition.Hidden = value ? MsoTriState.msoTrue : MsoTriState.msoFalse; }
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
            return FromSlideFactory(duplicatedSlide);
        }

        public bool HasAnimationForClick(int click)
        {
            var mainSequence = _slide.TimeLine.MainSequence;
            var effect = mainSequence.FindFirstAnimationForClick(click);

            return effect != null;
        }

        public void DeleteShapesWithPrefixTimelineInvariant(string prefix)
        {
            var mainSequence = _slide.TimeLine.MainSequence;
            var effectCnt = 1;

            while (effectCnt <= mainSequence.Count)
            {
                var effect = mainSequence[effectCnt];

                if (effect.Shape.Name.StartsWith(prefix))
                {
                    // if the shape is triggered on click, delete this may cause problem if the next
                    // effect is triggered with previous and we want the time sequence to be time
                    // invariant. To handle it, we need to set the on_prev event to be on_click.
                    if (effect.Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick &&
                        effect.Index + 1 <= mainSequence.Count)
                    {
                        var nextEffect = mainSequence[effect.Index + 1];

                        if (nextEffect.Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerWithPrevious)
                        {
                            nextEffect.Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                        }
                    }

                    effect.Delete();
                }
                else
                {
                    effectCnt++;
                }
            }

            DeleteShapesWithPrefix(prefix);
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

        public void DeleteShapeWithRule(Regex regex)
        {
            List<Shape> shapes = _slide.Shapes.Cast<Shape>().ToList();

            var matchingShapes = shapes.Where(current => regex.IsMatch(current.Name));
            foreach (Shape s in matchingShapes)
            {
                s.Delete();
            }
        }

        public void DeleteShapeWithName(string name)
        {
            List<Shape> shapes = _slide.Shapes.Cast<Shape>().ToList();
            var matchingShapes = shapes.Where(current => current.Name == name);

            foreach (var s in matchingShapes)
            {
                s.Delete();
            }
        }

        public void DeleteHiddenShapes()
        {
            _slide.Shapes
                .Cast<Shape>()
                .Where(sh => sh.Visible == MsoTriState.msoFalse)
                .ToList()
                .ForEach(sh => sh.Delete());
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

        public void SetShapeAsAutoplay(Shape shape)
        {
            var mainSequence = _slide.TimeLine.MainSequence;

            Effect firstClickEvent = mainSequence.FindFirstAnimationForClick(1);
            bool hasNoClicksOnSlide = firstClickEvent == null;

            if (hasNoClicksOnSlide)
            {
                AddShapeAsLastAutoplaying(shape, MsoAnimEffect.msoAnimEffectFade);
            }
            else
            {
                InsertAnimationBeforeExisting(shape, firstClickEvent, MsoAnimEffect.msoAnimEffectFade);
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

        public void InsertPicture(string fileName, MsoTriState linkToFile, MsoTriState saveWithDoc,
                                  Tuple<Single, Single> leftTopCorner)
        {
            _slide.Shapes.AddPicture(fileName, linkToFile, saveWithDoc, leftTopCorner.Item1, leftTopCorner.Item2).Select();
        }


        /// <summary>
        /// Creates a snapshot of snapshotSlide before entry animations and places an image of the slide in this slide
        /// Returns the image shape.
        /// </summary>
        public Shape InsertEntrySnapshotOfSlide(PowerPointSlide snapshotSlide)
        {
            PowerPointSlide nextSlideCopy = snapshotSlide.Duplicate();
            nextSlideCopy.Shapes
                            .Cast<Shape>()
                            .Where(shape => nextSlideCopy.HasEntryAnimation(shape))
                            .ToList()
                            .ForEach(shape => shape.Delete());

            nextSlideCopy.Copy();
            Shape slidePicture = _slide.Shapes.PasteSpecial(PpPasteDataType.ppPastePNG)[1];
            nextSlideCopy.Delete();
            return slidePicture;
        }


        /// <summary>
        /// Creates a snapshot of snapshotSlide after exit animations and places an image of the slide in this slide
        /// Returns the image shape.
        /// </summary>
        public Shape InsertExitSnapshotOfSlide(PowerPointSlide snapshotSlide)
        {
            PowerPointSlide previousSlideCopy = snapshotSlide.Duplicate();
            previousSlideCopy.Shapes
                            .Cast<Shape>()
                            .Where(shape => previousSlideCopy.HasExitAnimation(shape))
                            .ToList()
                            .ForEach(shape => shape.Delete());

            previousSlideCopy.Copy();
            Shape slidePicture = _slide.Shapes.PasteSpecial(PpPasteDataType.ppPastePNG)[1];
            previousSlideCopy.Delete();
            return slidePicture;
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
        
        /// <summary>
        /// Translates the x and y coordinates a VML Path String (obtained from MotionEffect.Path) by a specified amount.
        /// TODO: Not sure whether it works with any VML path string yet. Need to verify. It seems to. Idea: the numerical (non-alphabetical) values alternate between x and y coordinates. I translate every value this eay by either xShift or yShift.
        /// </summary>
        public string TranslateVmlPath(string path, float xShift, float yShift)
        {
            string[] splitPath = path.Split(' ');
            bool isXCoordinate = true;
            for (int i = 0; i < splitPath.Length; ++i)
            {
                string token = splitPath[i].Trim();
                if (token.Length <= 1 && char.IsLetter(token, 0)) continue;

                float val = float.Parse(token, CultureInfo.InvariantCulture);
                if (isXCoordinate)
                {
                    val += xShift;
                    isXCoordinate = false;
                }
                else
                {
                    val += yShift;
                    isXCoordinate = true;
                }
                splitPath[i] = val.ToString(CultureInfo.InvariantCulture);
            }
            return string.Join(" ", splitPath);
        }

        /// <summary>
        /// Changes the Left, Top coordinates and Width, Height of the shape while maintaining the positions of motion paths. 
        /// </summary>
        public void RelocateShapeWithoutPath(Shape shape, float newLeft, float newTop, float newWidth, float newHeight)
        {
            float originalLeft = shape.Left;
            float originalTop = shape.Top;
            float originalWidth = shape.Width;
            float originalHeight = shape.Height;
            shape.Left = newLeft;
            shape.Top = newTop;
            shape.Width = newWidth;
            shape.Height = newHeight;

            var effects = TimeLine.MainSequence.Cast<Effect>();
            // TODO: Generalize to paths other than msoAnimEffectPathDown?
            effects = effects.Where(e => e.Shape.Equals(shape) && e.EffectType == MsoAnimEffect.msoAnimEffectPathDown).ToList();

            float xShift = (originalLeft - newLeft) + (originalWidth - newWidth) / 2;
            float yShift = (originalTop - newTop) + (originalHeight - newHeight) / 2;
            xShift /= PowerPointPresentation.Current.SlideWidth;
            yShift /= PowerPointPresentation.Current.SlideHeight;

            foreach (var effect in effects)
            {
                var motionEffect = effect.Behaviors[1].MotionEffect;
                motionEffect.Path = TranslateVmlPath(motionEffect.Path, xShift, yShift);
            }
        }

        public Shape GroupShapes(IEnumerable<Shape> shapes)
        {
            return ToShapeRange(shapes).Group();
        }

        public ShapeRange ToShapeRange(IEnumerable<Shape> shapes)
        {
            var shapeList = shapes.ToList();
            var oldNames = shapeList.Select(shape => shape.Name).ToList();

            var currentShapeNames = Shapes.Cast<Shape>().Select(shape => shape.Name);
            var unusedNames = Common.GetUnusedStrings(currentShapeNames, shapeList.Count);
            shapeList.Zip(unusedNames, (shape, name) => shape.Name = name).ToList();


            var shapeRange = Shapes.Range(unusedNames);

            shapeList.Zip(oldNames, (shape, name) => shape.Name = name).ToList();

            return shapeRange;
        }

        /// <summary>
        /// Copies the shape into this slide, without the usual position offset when an existing shape is already there.
        /// </summary>
        public Shape CopyShapeToSlide(Shape shape)
        {
            try
            {
                shape.Copy();
                var newShape = _slide.Shapes.Paste()[1];
                newShape.Left = shape.Left;
                newShape.Top = shape.Top;
                return newShape;
            }
            catch (COMException)
            {
                // invalid shape for copy paste (e.g. a placeholder title box with no content)
                return null;
            }
        }

        /// <summary>
        /// Copies the shaperange into this slide, without the usual position offset when pasting over an existing shape.
        /// If you are having difficulty getting a shaperange, use the ToShapeRange method.
        /// TODO: Test this method more thoroughly in more cases other than Graphics.SquashSlides
        /// </summary>
        public ShapeRange CopyShapesToSlide(ShapeRange shapes)
        {
            // First Index all the shapes by name, so they can be identified later.
            int index = 0;
            var originalShapes = new Dictionary<string, Shape>();
            var originalNames = new Dictionary<string, string>();
            foreach (Shape shape in shapes)
            {
                var tempName = index.ToString();
                index++;

                originalNames.Add(tempName, shape.Name);
                originalShapes.Add(tempName, shape);
                // temporarily set the name before copy, so we can locate it again in the new slide.
                shape.Name = tempName;
            }

            // Copy all the shapes over.
            shapes.Copy();
            var newShapes = _slide.Shapes.Paste();

            // Now use the indexed names to set back the names and positions to the original shapes'
            foreach (Shape shape in newShapes)
            {
                var key = shape.Name;
                var originalName = originalNames[key];
                var originalShape = originalShapes[key];

                originalShape.Name = originalName;
                shape.Name = originalName;
                shape.Left = originalShape.Left;
                shape.Top = originalShape.Top;
            }

            return newShapes;
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

            Effect exitDetails = enumerableSequence.LastOrDefault(effect => effect.Shape.Equals(source));
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
            Sequence sequence = _slide.TimeLine.MainSequence;
            for (int x = sequence.Count; x >= 1; x--)
            {
                Effect effect = sequence[x];
                if (effect.Shape.Name == sh.Name && effect.Shape.Id == sh.Id)
                    effect.Delete();
            }
        }

        public void RemoveAnimationsForShapes(List<Shape> shapes)
        {
            foreach (Shape sh in shapes)
            {
                DeleteShapeAnimations(sh);
            }
        }

        public List<Shape> GetShapesWithPrefix(string prefix)
        {
            List<Shape> shapes = _slide.Shapes.Cast<Shape>().ToList();
            List<Shape> matchingShapes = shapes.Where(current => current.Name.StartsWith(prefix)).ToList();

            return matchingShapes;
        }

        public List<Shape> GetShapeWithName(string name)
        {
            List<Shape> shapes = _slide.Shapes.Cast<Shape>().ToList();
            List<Shape> matchingShapes = shapes.Where(current => current.Name == name).ToList();

            return matchingShapes;
        }

        /// <summary>
        /// Returns a dictionary shapeName => shape,
        /// where shape refers to the first (any) shape found in the slide with that name.
        /// </summary>
        public Dictionary<string, Shape> GetNameToShapeDictionary()
        {
            var shapes = _slide.Shapes.Cast<Shape>();
            var dictionary = new Dictionary<string, Shape>(shapes.Count());
            foreach (var shape in shapes)
            {
                if (!dictionary.ContainsKey(shape.Name))
                {
                    dictionary.Add(shape.Name, shape);
                }
            }
            return dictionary;
        }

        public Shape GetShape(Func<Shape, bool> condition)
        {
            return _slide.Shapes.Cast<Shape>().Where(condition).FirstOrDefault();
        }

        public List<Shape> GetShapesWithMediaType(PpMediaType type, Regex nameRule)
        {
            List<Shape> shapes = _slide.Shapes.Cast<Shape>().ToList();
            List<Shape> matchingShapes = shapes.Where(current => current.Type == MsoShapeType.msoMedia &&
                                                                 current.MediaType == type &&
                                                                 nameRule.IsMatch(current.Name)).ToList();

            return matchingShapes;
        }

        public List<Shape> GetShapesWithRule(Regex nameRule)
        {
            var shapes = _slide.Shapes.Cast<Shape>().ToList();
            var matchingShapes = shapes.Where(current => nameRule.IsMatch(current.Name)).ToList();

            return matchingShapes;
        }

        public List<Shape> GetShapesWithTypeAndRule(MsoShapeType type, Regex nameRule)
        {
            var shapes = _slide.Shapes.Cast<Shape>().ToList();
            var matchingShapes = shapes.Where(current => current.Type == type &&
                                              nameRule.IsMatch(current.Name)).ToList();

            return matchingShapes;
        }

        public bool HasShapeWithRule(Regex nameRule)
        {
            return GetShapesWithRule(nameRule).Count > 0;
        }

        public bool HasShapeWithSameName(string name)
        {
            List<Shape> shapes = _slide.Shapes.Cast<Shape>().ToList();
            List<Shape> matchingShapes = shapes.Where(current => current.Name == name).ToList();

            return matchingShapes.Count != 0;
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
            Slide duplicatedSlide = _slide.Duplicate()[1];
            return PowerPointDrillDownSlide.FromSlideFactory(duplicatedSlide);
        }

        public PowerPointSlide CreateStepBackSlide()
        {
            Slide duplicatedSlide = _slide.Duplicate()[1];
            return PowerPointStepBackSlide.FromSlideFactory(duplicatedSlide);
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
            Sequence sequence = _slide.TimeLine.MainSequence;
            for (int x = sequence.Count; x >= 1; x--)
            {
                Effect effect = sequence[x];
                if (effect.Shape.Name == shape.Name && effect.Shape.Id == shape.Id)
                    if (effect.Exit == Office.MsoTriState.msoTrue)
                        return true;
            }
            return false;
        }

        public bool HasEntryAnimation(Shape shape)
        {
            Sequence sequence = _slide.TimeLine.MainSequence;
            for (int x = sequence.Count; x >= 1; x--)
            {
                Effect effect = sequence[x];
                if (effect.Shape.Name == shape.Name && effect.Shape.Id == shape.Id)
                    if (IsEntryEffect(effect))
                        return true;
            }
            return false;
        }

        /// <summary>
        /// TODO: What does "Entry Animation" mean? entryEffects.Contains(effectType) could mean that it is either an entry or exit animation. Perhaps change it to entryEffects.Contains(effectType) && entryEffects.Exit == Mso False
        /// </summary>
        private bool IsEntryEffect(Effect effect)
        {
            return effect.Exit == MsoTriState.msoFalse && entryEffects.Contains(effect.EffectType);
        }

        /// <summary>
        /// Returns the index of the first effect in the slide that belongs to the shape.
        /// </summary>
        public int IndexOfFirstEffect(Shape shape)
        {
            Sequence sequence = _slide.TimeLine.MainSequence;
            for (int i = 1; i <= sequence.Count; i++)
            {
                Effect effect = sequence[i];
                if (effect.Shape.Name == shape.Name && effect.Shape.Id == shape.Id)
                    return i;
            }
            return -1;
        }

        public void DeletePlaceholderShapes()
        {
            _slide.Shapes.Placeholders.Cast<Shape>().ToList().ForEach(shape => shape.Delete());
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

        public Shape AddTemplateSlideMarker()
        {
            if (HasTemplateSlideMarker()) return null;

            float ratio = 22.5f;
            float slideWidth = PowerPointPresentation.Current.SlideWidth;
            float slideHeight = PowerPointPresentation.Current.SlideHeight;
            float shapeWidth = Math.Min(slideWidth, 900);
            float shapeHeight = shapeWidth/ratio;

            Shape markerShape = _slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, shapeWidth, shapeHeight);

            markerShape.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentCentered;

            markerShape.TextFrame2.TextRange.Text = TextCollection.AgendaLabTemplateSlideInstructions;
            markerShape.Fill.ForeColor.RGB = 0x0000C0;
            markerShape.TextFrame2.TextRange.Font.Bold = MsoTriState.msoTrue;
            markerShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0x00FFFF;
            markerShape.TextFrame2.TextRange.Paragraphs[2].Font.Fill.ForeColor.RGB = 0xFFFFFF;
            markerShape.TextFrame2.TextRange.Paragraphs[2].Font.Bold = MsoTriState.msoFalse;

            markerShape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
            markerShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            
            markerShape.Left = (slideWidth - markerShape.Width) / 2;
            markerShape.Top = slideHeight - markerShape.Height;
            markerShape.Name = PptLabsTemplateMarkerShapeName;

            Utils.Graphics.MakeShapeViewTimeInvisible(markerShape, _slide);
            return markerShape;
        }

        public bool HasTemplateSlideMarker()
        {
            return _slide.Shapes.Cast<Shape>().Any(IsTemplateSlideMarker);
        }

        public static bool IsTemplateSlideMarker(Shape shape)
        {
            return shape.Name == PptLabsTemplateMarkerShapeName;
        }

        public static bool IsNotTemplateSlideMarker(Shape shape)
        {
            return !IsTemplateSlideMarker(shape);
        }

        protected Shape AddPowerPointLabsIndicator()
        {
            String tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            Shape indicatorShape = _slide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, PowerPointPresentation.Current.SlideWidth - 120, 0, 120, 84);

            indicatorShape.Left = PowerPointPresentation.Current.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.Name = PptLabsIndicatorShapeName + DateTime.Now.ToString("yyyyMMddHHmmssffff");

            Utils.Graphics.MakeShapeViewTimeInvisible(indicatorShape, _slide);

            return indicatorShape;
        }

        public void DeleteIndicator()
        {
            _slide.Shapes.Cast<Shape>()
                        .Where(IsIndicator)
                        .ToList()
                        .ForEach(shape => shape.Delete());
        }


        public void HideIndicator()
        {
            _slide.Shapes.Cast<Shape>()
                        .Where(IsIndicator)
                        .ToList()
                        .ForEach(shape => shape.Visible = MsoTriState.msoFalse);
        }

        public void ShowIndicator()
        {
            _slide.Shapes.Cast<Shape>()
                        .Where(IsIndicator)
                        .ToList()
                        .ForEach(shape => shape.Visible = MsoTriState.msoTrue);
        }

        public void BringIndicatorToFront()
        {
            _slide.Shapes.Cast<Shape>()
                        .Where(IsIndicator)
                        .ToList()
                        .ForEach(shape => shape.ZOrder(MsoZOrderCmd.msoBringToFront));
        }

        public static bool IsIndicator(Shape shape)
        {
            return shape.Name.StartsWith(PptLabsIndicatorShapeName);
        }

        protected void RemoveSlideTransitions()
        {
            _slide.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectNone;
        }

        public void MoveMotionAnimation()
        {
            Sequence sequence = _slide.TimeLine.MainSequence;
            foreach (Effect eff in _slide.TimeLine.MainSequence)
            {
                if ((eff.EffectType >= MsoAnimEffect.msoAnimEffectPathCircle && eff.EffectType <= MsoAnimEffect.msoAnimEffectPathRight) || eff.EffectType == MsoAnimEffect.msoAnimEffectCustom)
                {
                    AnimationBehavior motion = eff.Behaviors[1];
                    if (motion.Type == MsoAnimType.msoAnimTypeMotion)
                    {
                        Shape sh = eff.Shape;
                        string motionPath = motion.MotionEffect.Path.Trim();
                        if (motionPath.Last() < 'A' || motionPath.Last() > 'Z')
                            motionPath += " X";
                        string[] path = motionPath.Split(' ');
                        int count = path.Length;
                        float xVal = Convert.ToSingle(path[count - 3]);
                        float yVal = Convert.ToSingle(path[count - 2]);
                        sh.Left += (xVal * PowerPointPresentation.Current.SlideWidth);
                        sh.Top += (yVal * PowerPointPresentation.Current.SlideHeight);
                    }
                }
            }
        }

        public void AddAppearDisappearAnimation(Shape sh)
        {
            Sequence sequence = _slide.TimeLine.MainSequence;
            Effect effectAppear = sequence.AddEffect(sh, MsoAnimEffect.msoAnimEffectAppear, MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectAppear.Timing.Duration = 0;

            Effect effectDisappear = sequence.AddEffect(sh, MsoAnimEffect.msoAnimEffectAppear, MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;
        }

        public Shape GetShapeWithSameIDAndName(Shape shapeToMatch)
        {
            Shape tempMatchingShape = null;
            foreach (Shape sh in _slide.Shapes)
            {
                if (shapeToMatch.Id == sh.Id && HaveSameNames(shapeToMatch, sh))
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
                if (HaveSameNames(shapeToMatch, sh))
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

        public bool IsSpotlightSlide()
        {
            return _slide.Name.Contains("PPTLabsSpotlight");
        }

        public bool IsAckSlide()
        {
            return PowerPointAckSlide.IsAckSlide(this);
        }

        public PowerPointSlide CreateAckSlide()
        {
            Slide ackSlide = PowerPointPresentation.Current.Presentation.Slides.Add(PowerPointPresentation.Current.SlideCount + 1, PpSlideLayout.ppLayoutBlank);
            return PowerPointAckSlide.FromSlideFactory(ackSlide);
        }

        public bool HasTextFragments()
        {
            foreach (Shape sh in _slide.Shapes)
            {
                if (sh.Name.StartsWith("PPTLabsHighlightTextFragmentsShape"))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Get all shapes in the slide, ordered by when they appear in the animation timeline.
        /// Shapes with entry animations are ordered behind shapes without entry animations.
        /// Shapes without animations are placed at the back.
        /// </summary>
        public List<Shape> GetShapesOrderedByTimeline()
        {
            var shapesWithEntry = new List<Shape>();
            var shapesWithoutEntry = new List<Shape>();
            var identifiedShapeIds = new HashSet<int>();

            var seq = _slide.TimeLine.MainSequence;
            for (int i = 1; i <= seq.Count; ++i)
            {
                var effect = seq[i];
                var shape = effect.Shape;
                if (!identifiedShapeIds.Contains(shape.Id))
                {
                    identifiedShapeIds.Add(shape.Id);
                    if (IsEntryEffect(effect))
                    {
                        shapesWithEntry.Add(shape);
                    }
                    else
                    {
                        shapesWithoutEntry.Add(shape);
                    }
                }
            }

            var remainingShapes = _slide.Shapes.Cast<Shape>().Where(shape => !identifiedShapeIds.Contains(shape.Id));

            var shapes = shapesWithoutEntry;
            shapes.AddRange(shapesWithEntry);
            shapes.AddRange(remainingShapes);
            return shapes;
        }

        /// <summary>
        /// Returns all HighlightTextFragmentsShapes in the order they appear in the animation timeline.
        /// </summary>
        public List<Shape> GetTextFragments()
        {
            var allShapes = GetShapesOrderedByTimeline();
            return allShapes.Where(shape => shape.Name.StartsWith("PPTLabsHighlightTextFragmentsShape")).ToList();
        }

        public bool HasCaptions()
        {
            foreach (Shape shape in this.Shapes)
            {
                if (shape.Name.StartsWith("PowerPointLabs Caption"))
                {
                    return true;
                }
            }
            return false;
        }

        public bool HasAudio()
        {
            foreach (Shape shape in this.Shapes)
            {
                if (shape.Name.Contains(NotesToAudio.SpeechShapePrefix) || 
                    shape.Name.Contains(NotesToAudio.SpeechShapePrefixOld))
                {
                    return true;
                }
            }
            return false;
        }

        public Effect AddShapeAsLastAutoplaying(Shape shape, MsoAnimEffect effect)
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

        private bool IsNextSlideTransitionBlacklisted()
        {
            bool isLastSlide = _slide.SlideIndex == PowerPointPresentation.Current.SlideCount;
            if (isLastSlide)
            {
                return false;
            }

            // Indexes are from 1, while the slide collection starts from 0.
            PowerPointSlide nextSlide = PowerPointPresentation.Current.Slides[Index];
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

        private static void DeleteEffectsForShape(Shape shape, IEnumerable<Effect> mainEffects)
        {
            var shapeToDeleteList = mainEffects.Where(e => e.Shape.Equals(shape)).ToList();
            
            foreach (Effect e in shapeToDeleteList)
            {
                e.Delete();
            }
        }

        private float GetDistanceBetweenShapes(Shape sh1, Shape sh2)
        {
            float sh1CenterX = (sh1.Left + (sh1.Width / 2));
            float sh2CenterX = (sh2.Left + (sh2.Width / 2));
            float sh1CenterY = (sh1.Top + (sh1.Height / 2));
            float sh2CenterY = (sh2.Top + (sh2.Height / 2));
            float distSquared = (float)(Math.Pow((sh2CenterX - sh1CenterX), 2) + Math.Pow((sh2CenterY - sh1CenterY), 2));
            return (float)(Math.Sqrt(distSquared));
        }

        private bool HaveSameNames(Shape sh1, Shape sh2)
        {
            String name1 = sh1.Name;
            String name2 = sh2.Name;

            return (name1.ToUpper().CompareTo(name2.ToUpper()) == 0);
        }

        /// <summary>
        /// Default shapes have the property where if you duplicate them (or copy/paste), they change names.
        /// This command renames the shapes in the slide so that they don't have the default names.
        /// </summary>
        public void MakeShapeNamesNonDefault()
        {
            var shapes = _slide.Shapes.Cast<Shape>();
            foreach (var shape in shapes)
            {
                if (Graphics.HasDefaultName(shape))
                {
                    shape.Name = UnnamedShapeName + Common.UniqueDigitString();
                }
            }
        }

        /// <summary>
        /// Gives all shapes in the slide unique names. Good to call before sync logic.
        /// Note: If the name of the shape is used to identify the shape (e.g. through AgendaShape),
        /// this can be dangerous if there are duplicates as it overrides the original name.
        /// </summary>
        public void MakeShapeNamesUnique(Func<Shape, bool> restrictTo = null)
        {
            if (restrictTo == null) restrictTo = shape => true;

            var currentNames = new HashSet<string>();
            var shapes = _slide.Shapes.Cast<Shape>().Where(restrictTo);

            foreach (var shape in shapes)
            {
                if (currentNames.Contains(shape.Name))
                {
                    shape.Name = UnnamedShapeName + Common.UniqueDigitString();
                }
                currentNames.Add(shape.Name);
            }
        }

        public void DeleteSlideNumberShapes()
        {
            List<Shape> shapes = _slide.Shapes.Cast<Shape>().ToList();

            var matchingShapes = shapes.Where(current => current.Type == MsoShapeType.msoPlaceholder && current.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderSlideNumber);

            foreach (Shape s in matchingShapes)
            {
                s.Delete();
            }
        }
    }
}
