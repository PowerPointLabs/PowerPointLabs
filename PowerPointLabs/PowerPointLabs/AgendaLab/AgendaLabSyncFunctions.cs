using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using Graphics = PowerPointLabs.Utils.Graphics;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.AgendaLab
{
    internal static partial class AgendaLabMain
    {
        #region Bullet Agenda

        /// <summary>
        /// The SyncFunction used for synchronising the bullet agenda slides.
        /// </summary>
        public static readonly SyncFunction SyncBulletAgendaSlide = (refSlide, sections, currentSection, targetSlide) =>
        {
            SyncShapesFromReferenceSlide(refSlide, targetSlide);

            targetSlide.Transition.EntryEffect = PpEntryEffect.ppEffectFadeSmoothly;
            targetSlide.Transition.Duration = 0.25f;

            var referenceContentShape = refSlide.GetShape(AgendaShape.WithPurpose(ShapePurpose.ContentShape));
            var targetContentShape = targetSlide.GetShape(AgendaShape.WithPurpose(ShapePurpose.ContentShape));
            var bulletFormats = BulletFormats.ExtractFormats(referenceContentShape);

            Graphics.SetText(targetContentShape, sections.Where(section => section.Index > 1)
                                                        .Select(section => section.Name));
            Graphics.SyncShape(referenceContentShape, targetContentShape, pickupTextContent: false,
                pickupTextFormat: false);

            ApplyBulletFormats(targetContentShape.TextFrame2.TextRange, bulletFormats, currentSection);
            targetSlide.DeletePlaceholderShapes();
        };


        private static void ApplyBulletFormats(TextRange2 textRange, BulletFormats bulletFormats, AgendaSection currentSection)
        {
            // - 1 because first section in agenda is at index 2 (exclude first section)
            int focusIndex = currentSection.Index - 1;

            for (var i = 1; i <= textRange.Paragraphs.Count; i++)
            {
                var currentParagraph = textRange.Paragraphs[i];

                if (i == focusIndex)
                {
                    Graphics.SyncTextRange(bulletFormats.Highlighted, currentParagraph, pickupTextContent: false);
                }
                else if (i < focusIndex)
                {
                    Graphics.SyncTextRange(bulletFormats.Visited, currentParagraph, pickupTextContent: false);
                }
                else
                {
                    Graphics.SyncTextRange(bulletFormats.Unvisited, currentParagraph, pickupTextContent: false);
                }
            }
        }

        #endregion


        #region Visual Agenda

        /// <summary>
        /// The SyncFunction used for synchronising the Visual agenda slides (other than the last visual agenda slide).
        /// </summary>
        public static readonly SyncFunction SyncVisualAgendaSlide = (refSlide, sections, currentSection, targetSlide) =>
        {
            DeleteVisualAgendaImageShapes(targetSlide);
            SyncShapesFromReferenceSlide(refSlide, targetSlide);
            ReplaceVisualImagesWithAfterZoomOutImages(targetSlide, currentSection.Index);

            if (currentSection.Index > 2)
            {
                // Not first visual section.
                var zoomOutShape = FindShapeCorrespondingToSection(targetSlide, currentSection.Index - 1);
                GenerateVisualAgendaSlideZoomOut(targetSlide, zoomOutShape);
            }

            var zoomInShape = FindShapeCorrespondingToSection(targetSlide, currentSection.Index);
            GenerateVisualAgendaSlideZoomIn(targetSlide, zoomInShape);

            targetSlide.DeletePlaceholderShapes();
        };

        /// <summary>
        /// The SyncFunction used for synchronising the last visual agenda slide.
        /// </summary>
        public static readonly SyncFunction SyncVisualAgendaEndSlide = (refSlide, sections, currentSection, targetSlide) =>
        {
            DeleteVisualAgendaImageShapes(targetSlide);
            SyncShapesFromReferenceSlide(refSlide, targetSlide);
            ReplaceVisualImagesWithAfterZoomOutImages(targetSlide, currentSection.Index + 1);

            var zoomOutShape = FindShapeCorrespondingToSection(targetSlide, currentSection.Index);
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
            var indexedShapes = new Dictionary<int, Shape>();
            slide.Shapes.Cast<Shape>()
                        .Where(AgendaShape.WithPurpose(ShapePurpose.VisualAgendaImage))
                        .ToList()
                        .ForEach(shape => indexedShapes.Add(AgendaShape.Decode(shape).Section.Index, shape));

            for (int i = 2; i < sectionIndex; ++i)
            {
                var imageShape = indexedShapes[i];

                var sectionEndSlide = FindSectionLastNonAgendaSlide(i);
                var snapshotShape = slide.InsertExitSnapshotOfSlide(sectionEndSlide);
                snapshotShape.Name = imageShape.Name;
                Graphics.SyncShape(imageShape, snapshotShape, pickupShapeFormat: false, pickupTextContent: false, pickupTextFormat: false);
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

            var index = slide.Index;

            // move the step back slide to the first slide of the section
            PowerPointPresentation.Current.Presentation.Slides[index - 1].MoveTo(index);
            slide.MoveTo(index);
        }
 
        #endregion


        #region General

        /// <summary>
        /// Synchronises the shapes in the candidate slide with the shapes in the reference slide.
        /// Adds any shape that exists in the reference slide but is missing in the candidate slide.
        /// </summary>
        private static void SyncShapesFromReferenceSlide(PowerPointSlide refSlide, PowerPointSlide candidate)
        {
            if (refSlide == null || candidate == null || refSlide == candidate)
            {
                return;
            }

            refSlide.MakeShapeNamesNonDefault();
            refSlide.MakeShapeNamesUnique(shape => !AgendaShape.IsAnyAgendaShape(shape) &&
                                                   !PowerPointSlide.IsTemplateSlideMarker(shape));

            candidate.Layout = refSlide.Layout;
            candidate.Design = refSlide.Design;

            // syncronize extra shapes other than visual items in reference slide
            var extraShapes = refSlide.Shapes.Cast<Shape>()
                                             .Where(shape => !candidate.HasShapeWithSameName(shape.Name) &&
                                                             !PowerPointSlide.IsIndicator(shape) &&
                                                             !PowerPointSlide.IsTemplateSlideMarker(shape))
                                             .Select(shape => shape.Name)
                                             .ToArray();

            if (extraShapes.Length != 0)
            {
                var refShapes = refSlide.Shapes.Range(extraShapes);
                refShapes.Copy();
                var copiedShapes = candidate.Shapes.Paste();
            }

            // syncronize shapes position and size, except bullet content
            var candidateSlideShapes = candidate.GetNameToShapeDictionary();
            var sameShapes = refSlide.Shapes.Cast<Shape>()
                                            .Where(shape => !PowerPointSlide.IsIndicator(shape) &&
                                                            !PowerPointSlide.IsTemplateSlideMarker(shape) &&
                                                            candidateSlideShapes.ContainsKey(shape.Name));

            var shapeOriginalZOrders = new SortedDictionary<int, Shape>();
            foreach (var refShape in sameShapes)
            {
                var candidateShape = candidateSlideShapes[refShape.Name];
                Graphics.SyncShape(refShape, candidateShape);

                shapeOriginalZOrders.Add(refShape.ZOrderPosition, candidateShape);
            }

            SynchroniseZOrders(shapeOriginalZOrders);
        }

        private static void SynchroniseZOrders(SortedDictionary<int, Shape> shapeOriginalZOrders)
        {
            Shape lastShape = null;
            foreach (var entry in shapeOriginalZOrders.Reverse())
            {
                var shape = entry.Value;
                if (lastShape != null)
                {
                    Graphics.MoveZUntilInFront(shape, lastShape);
                }
                lastShape = shape;
            }
        }

        #endregion
    }
}
