using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.AgendaLab.Templates;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using PowerPointLabs.ZoomLab;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

namespace PowerPointLabs.AgendaLab
{
    internal static partial class AgendaLabMain
    {
#pragma warning disable 0618
        // This file contains Sync Functions, which are used to sync individual slides (not the agenda as a whole).
        // The methods defined in this file are helper methods for the sync functions.

        #region Bullet Agenda

        /// <summary>
        /// The SyncFunction used for synchronizing the front bullet agenda slides.
        /// </summary>
        public static readonly SyncFunction SyncStartBulletAgendaSlide = (refSlide, sections, currentSection, deletedShapeNames, isNewlyGenerated, targetSlide) =>
        {
            SyncBulletAgendaSlide(refSlide, sections, currentSection, deletedShapeNames, targetSlide);

            if (isNewlyGenerated)
            {
                targetSlide.Transition.EntryEffect = PpEntryEffect.ppEffectFadeSmoothly;
                targetSlide.Transition.Duration = 0.25f;

                PowerPointSlide nextSlide = TryGetSlideAtIndex(targetSlide.Index + 1);
                if (nextSlide != null)
                {
                    nextSlide.Transition.EntryEffect = refSlide.Transition.EntryEffect;
                    nextSlide.Transition.Duration = refSlide.Transition.Duration;
                }
            }
        };

        /// <summary>
        /// The SyncFunction used for synchronizing the end bullet agenda slides.
        /// </summary>
        public static readonly SyncFunction SyncEndBulletAgendaSlide = (refSlide, sections, currentSection, deletedShapeNames, isNewlyGenerated, targetSlide) =>
        {
            SyncBulletAgendaSlide(refSlide, sections, currentSection, deletedShapeNames, targetSlide);

            if (isNewlyGenerated)
            {
                targetSlide.Transition.EntryEffect = refSlide.Transition.EntryEffect;
                targetSlide.Transition.Duration = refSlide.Transition.Duration;
            }
        };

        /// <summary>
        /// The SyncFunction used for synchronizing the final bullet agenda slide.
        /// </summary>
        public static readonly SyncFunction SyncFinalBulletAgendaSlide = (refSlide, sections, currentSection, deletedShapeNames, isNewlyGenerated, targetSlide) =>
        {
            SyncStartBulletAgendaSlide(refSlide, sections, AgendaSection.None, deletedShapeNames, isNewlyGenerated, targetSlide);
        };

        private static void SyncBulletAgendaSlide(PowerPointSlide refSlide, List<AgendaSection> sections,
            AgendaSection currentSection, List<string> deletedShapeNames, PowerPointSlide targetSlide)
        {
            SyncShapesFromReferenceSlide(refSlide, targetSlide, deletedShapeNames);

            Shape referenceContentShape = refSlide.GetShape(AgendaShape.WithPurpose(ShapePurpose.ContentShape));
            Shape targetContentShape = targetSlide.GetShape(AgendaShape.WithPurpose(ShapePurpose.ContentShape));
            BulletFormats bulletFormats = BulletFormats.ExtractFormats(referenceContentShape);

            ShapeUtil.SetText(targetContentShape, sections.Where(section => section.Index > 1)
                .Select(section => section.Name));
            ShapeUtil.SyncShape(referenceContentShape, targetContentShape, pickupTextContent: false,
                pickupTextFormat: false);

            ApplyBulletFormats(targetContentShape.TextFrame2.TextRange, bulletFormats, currentSection);
            targetSlide.DeletePlaceholderShapes();
        }

        /// <summary>
        /// Applies font highlighting by section to the text in the bullet agenda.
        /// Set currentSection to the first section for everything to be unvisited.
        /// Set currentSection to AgendaSection.None for everything to be visited.
        /// </summary>
        private static void ApplyBulletFormats(TextRange2 textRange, BulletFormats bulletFormats, AgendaSection currentSection)
        {
            // - 1 because first section in agenda is at index 2 (exclude first section)
            int focusIndex = currentSection.IsNone() ? int.MaxValue : currentSection.Index - 1;

            textRange.Font.StrikeThrough = MsoTriState.msoFalse;

            for (int i = 1; i <= textRange.Paragraphs.Count; i++)
            {
                TextRange2 currentParagraph = textRange.Paragraphs[i];

                if (i == focusIndex)
                {
                    ShapeUtil.SyncTextRange(bulletFormats.Highlighted, currentParagraph, pickupTextContent: false);
                }
                else if (i < focusIndex)
                {
                    ShapeUtil.SyncTextRange(bulletFormats.Visited, currentParagraph, pickupTextContent: false);
                }
                else
                {
                    ShapeUtil.SyncTextRange(bulletFormats.Unvisited, currentParagraph, pickupTextContent: false);
                }
            }
        }

        #endregion


        #region Visual Agenda

        /// <summary>
        /// The SyncFunction used for synchronizing the Visual agenda slides (other than the last visual agenda slide).
        /// </summary>
        public static readonly SyncFunction SyncVisualAgendaSlide = (refSlide, sections, currentSection, deletedShapeNames, isNewlyGenerated, targetSlide) =>
        {
            DeleteVisualAgendaImageShapes(targetSlide);
            SyncShapesFromReferenceSlide(refSlide, targetSlide, deletedShapeNames);
            ReplaceVisualImagesWithAfterZoomOutImages(targetSlide, currentSection.Index);

            if (currentSection.Index > 2)
            {
                // Not first visual section.
                Shape zoomOutShape = FindShapeCorrespondingToSection(targetSlide, currentSection.Index - 1);
                GenerateVisualAgendaSlideZoomOut(targetSlide, zoomOutShape);
            }

            Shape zoomInShape = FindShapeCorrespondingToSection(targetSlide, currentSection.Index);
            GenerateVisualAgendaSlideZoomIn(targetSlide, zoomInShape);

            targetSlide.DeletePlaceholderShapes();
        };

        /// <summary>
        /// The SyncFunction used for synchronizing the last visual agenda slide.
        /// </summary>
        public static readonly SyncFunction SyncVisualAgendaEndSlide = (refSlide, sections, currentSection, deletedShapeNames, isNewlyGenerated, targetSlide) =>
        {
            DeleteVisualAgendaImageShapes(targetSlide);
            SyncShapesFromReferenceSlide(refSlide, targetSlide, deletedShapeNames);
            ReplaceVisualImagesWithAfterZoomOutImages(targetSlide, currentSection.Index + 1);

            Shape zoomOutShape = FindShapeCorrespondingToSection(targetSlide, currentSection.Index);
            GenerateVisualAgendaSlideZoomOut(targetSlide, zoomOutShape, finalZoomOut: true);

            targetSlide.DeletePlaceholderShapes();
        };

        private static void DeleteVisualAgendaImageShapes(PowerPointSlide slide)
        {
            slide.Shapes.Cast<Shape>()
                .Where(AgendaShape.WithPurpose(ShapePurpose.VisualAgendaImage))
                .ToList()
                .ForEach(shape => shape.Delete());
        }

        /// <summary>
        /// Within the slide, for all sections that have been "passed", replace their visual agenda image shape with
        /// an image of the end slide of the section.
        /// </summary>
        private static void ReplaceVisualImagesWithAfterZoomOutImages(PowerPointSlide slide, int sectionIndex)
        {
            Dictionary<int, Shape> indexedShapes = new Dictionary<int, Shape>();
            slide.Shapes.Cast<Shape>()
                        .Where(AgendaShape.WithPurpose(ShapePurpose.VisualAgendaImage))
                        .ToList()
                        .ForEach(shape => indexedShapes.Add(AgendaShape.Decode(shape).Section.Index, shape));

            for (int i = 2; i < sectionIndex; ++i)
            {
                Shape imageShape = indexedShapes[i];

                PowerPointSlide sectionEndSlide = FindSectionLastNonAgendaSlide(i);
                Shape snapshotShape = slide.InsertExitSnapshotOfSlide(sectionEndSlide);
                snapshotShape.Name = imageShape.Name;
                ShapeUtil.SyncShape(imageShape, snapshotShape, pickupShapeFormat: true, pickupTextContent: false, pickupTextFormat: false);
                imageShape.Delete();
            }
        }

        /// <summary>
        /// Searches for the visual agenda image shape that corresponds to the given section index in the slide and returns it.
        /// </summary>
        private static Shape FindShapeCorrespondingToSection(PowerPointSlide inSlide, int sectionIndex)
        {
            return inSlide.GetShape(AgendaShape.MeetsConditions(shape => shape.ShapePurpose == ShapePurpose.VisualAgendaImage &&
                                                                        sectionIndex == shape.Section.Index));
        }

        /// <summary>
        /// Create the zoom in (drill down) effect in visual agenda. The zoom in slide is not part of the template.
        /// </summary>
        private static void GenerateVisualAgendaSlideZoomIn(PowerPointSlide slide, Shape zoomInShape)
        {
            PowerPointDrillDownSlide addedSlide;
            AutoZoom.AddDrillDownAnimation(zoomInShape, slide, out addedSlide, includeAckSlide: false, deletePreviouslyAdded: false);
            slide.GetShapesWithRule(new Regex("PPTZoomIn"))[0].Delete();
            AgendaSection section = AgendaSlide.Decode(slide).Section;
            AgendaSlide.SetSlideName(addedSlide, Type.Visual, SlidePurpose.ZoomIn, section);
            zoomInShape.Visible = MsoTriState.msoTrue;
        }

        /// <summary>
        /// Create the zoom out (step back) effect in visual agenda. The zoom out slide is not part of the template.
        /// </summary>
        private static void GenerateVisualAgendaSlideZoomOut(PowerPointSlide slide, Shape zoomOutShape, bool finalZoomOut = false)
        {
            PowerPointStepBackSlide addedSlide;
            AutoZoom.AddStepBackAnimation(zoomOutShape, slide, out addedSlide, includeAckSlide: false, deletePreviouslyAdded: false);
            slide.GetShapesWithRule(new Regex("PPTZoomOut"))[0].Delete();
            AgendaSection section = AgendaSlide.Decode(slide).Section;
            AgendaSlide.SetSlideName(addedSlide, Type.Visual, finalZoomOut ? SlidePurpose.FinalZoomOut : SlidePurpose.ZoomOut, section);
            zoomOutShape.Visible = MsoTriState.msoTrue;

            int index = slide.Index;

            // move the step back slide to the first slide of the section
            PowerPointPresentation.Current.Presentation.Slides[index - 1].MoveTo(index);
            slide.MoveTo(index);
        }
 
        #endregion


        #region General

        /// <summary>
        /// Synchronizes the shapes in the candidate slide with the shapes in the reference slide.
        /// Adds any shape that exists in the reference slide but is missing in the candidate slide.
        /// </summary>
        private static void SyncShapesFromReferenceSlide(PowerPointSlide refSlide, PowerPointSlide candidate, List<string> markedForDeletion)
        {
            if (refSlide == null || candidate == null || refSlide == candidate)
            {
                return;
            }

            DeleteShapesMarkedForDeletion(candidate, markedForDeletion);

            candidate.Layout = refSlide.Layout;
            
            candidate.Design = SlideUtil.GetDesign("Agenda Template");
            candidate.GetNativeSlide().FollowMasterBackground = MsoTriState.msoTrue;

            // synchronize extra shapes other than visual items in reference slide
            Dictionary<string, Shape> candidateSlideShapes = candidate.GetNameToShapeDictionary();
            string[] extraShapes = refSlide.Shapes.Cast<Shape>()
                                             .Where(shape => !PowerPointSlide.IsIndicator(shape) &&
                                                             !PowerPointSlide.IsTemplateSlideMarker(shape) &&
                                                             !candidateSlideShapes.ContainsKey(shape.Name))
                                             .Select(shape => shape.Name)
                                             .ToArray();

            if (extraShapes.Length != 0)
            {
                ShapeRange refShapes = refSlide.Shapes.Range(extraShapes);
                CopyShapesTo(refShapes, candidate);
            }

            // synchronize shapes position and size, except bullet content
            candidateSlideShapes = candidate.GetNameToShapeDictionary();
            IEnumerable<Shape> sameShapes = refSlide.Shapes.Cast<Shape>()
                                            .Where(shape => !PowerPointSlide.IsIndicator(shape) &&
                                                            !PowerPointSlide.IsTemplateSlideMarker(shape) &&
                                                            candidateSlideShapes.ContainsKey(shape.Name));

            SortedDictionary<int, Shape> shapeOriginalZOrders = new SortedDictionary<int, Shape>();
            foreach (Shape refShape in sameShapes)
            {
                Shape candidateShape = candidateSlideShapes[refShape.Name];
                ShapeUtil.SyncWholeShape(refShape, ref candidateShape, candidate);

                shapeOriginalZOrders.Add(refShape.ZOrderPosition, candidateShape);
            }

            SynchronizeZOrders(shapeOriginalZOrders);
        }

        private static void CopyShapesTo(ShapeRange refShapes, PowerPointSlide candidate)
        {
            foreach (Shape shape in refShapes)
            {
                try
                {
                    shape.Copy();
                    candidate.Shapes.Paste();
                }
                catch (COMException)
                {
                    // A COMException occurs if you try to copy paste an empty placeholder shape. So I catch it here.
                    // I can't figure out any other way to detect that it's an empty placeholder shape.
                    // You know, those things like "Click to add title..."
                }
            }
        }

        private static void DeleteShapesMarkedForDeletion(PowerPointSlide candidate, List<string> markedForDeletion)
        {
            if (markedForDeletion.Count == 0)
            {
                return;
            }
            
            Dictionary<string, Shape> candidateSlideShapes = candidate.GetNameToShapeDictionary();
            foreach (string shapeName in markedForDeletion)
            {
                Shape shapeInSlide;
                bool shapeExists = candidateSlideShapes.TryGetValue(shapeName, out shapeInSlide);
                if (!shapeExists || shapeInSlide == null)
                {
                    continue;
                }
                
                shapeInSlide.Delete();
                candidateSlideShapes[shapeName] = null;
            }
        }

        private static void SynchronizeZOrders(SortedDictionary<int, Shape> shapeOriginalZOrders)
        {
            Shape lastShape = null;
            foreach (KeyValuePair<int, Shape> entry in shapeOriginalZOrders.Reverse())
            {
                Shape shape = entry.Value;
                if (lastShape != null)
                {
                    ShapeUtil.MoveZUntilBehind(shape, lastShape);
                }
                lastShape = shape;
            }
        }

        #endregion
    }
}
