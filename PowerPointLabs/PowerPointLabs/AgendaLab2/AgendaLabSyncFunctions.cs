using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;
using Graphics = PowerPointLabs.Utils.Graphics;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.AgendaLab2
{
    internal static partial class AgendaLabMain
    {
        #region Bullet Agenda

        public static readonly SyncFunction SyncBulletAgendaSlide = (refSlide, sections, currentSection, targetSlide) =>
        {
            AdjustBulletTemplateContent(refSlide, sections.Count);
            SyncSingleAgendaGeneral(refSlide, targetSlide);

            var referenceContentShape = refSlide.GetShape(AgendaShape.WithPurpose(ShapePurpose.ContentShape));
            var targetContentShape = targetSlide.GetShape(AgendaShape.WithPurpose(ShapePurpose.ContentShape));
            var bulletFormats = BulletFormats.ExtractFormats(referenceContentShape);

            Graphics.SetText(targetContentShape, sections.Where(section => section.Index > 1)
                                                        .Select(section => section.Name));
            Graphics.SyncShape(referenceContentShape, targetContentShape, pickupTextContent: false,
                pickupTextFormat: false);

            ReformatTextRange(targetContentShape.TextFrame2.TextRange, bulletFormats, currentSection);
            targetSlide.DeletePlaceholderShapes();
        };


        private static void AdjustBulletTemplateContent(PowerPointSlide refSlide, int numberOfSections)
        {
            // post process bullet points
            var contentHolder = refSlide.GetShape(AgendaShape.WithPurpose(ShapePurpose.ContentShape));
            var textRange = contentHolder.TextFrame2.TextRange;

            while (textRange.Paragraphs.Count < numberOfSections)
            {
                textRange.InsertAfter("\r ");
            }

            while (textRange.Paragraphs.Count > 3 && textRange.Paragraphs.Count > numberOfSections)
            {
                textRange.Paragraphs[textRange.Paragraphs.Count].Delete();
            }

            for (var i = 4; i <= textRange.Paragraphs.Count; i++)
            {
                textRange.Paragraphs[i].ParagraphFormat.Bullet.Type = MsoBulletType.msoBulletNone;
            }
        }

        private static void ReformatTextRange(TextRange2 textRange, BulletFormats bulletFormats, AgendaSection currentSection)
        {
            // - 1 because first section in agenda is at index 2 (exclude first section)
            int focusIndex = currentSection.Index - 1;

            for (var i = 1; i <= textRange.Paragraphs.Count; i++)
            {
                var curPara = textRange.Paragraphs[i];

                if (i == focusIndex)
                {
                    Graphics.SyncTextRange(bulletFormats.Highlighted, curPara, pickupTextContent: false);
                }
                else if (i < focusIndex)
                {
                    Graphics.SyncTextRange(bulletFormats.Visited, curPara, pickupTextContent: false);
                }
                else
                {
                    Graphics.SyncTextRange(bulletFormats.Unvisited, curPara, pickupTextContent: false);
                }
            }
        }

        #endregion


        #region Visual Agenda

        public static readonly SyncFunction SyncVisualAgendaSlide = (refSlide, sections, currentSection, targetSlide) =>
        {
            DeleteVisualAgendaImageShapes(targetSlide);
            SyncSingleAgendaGeneral(refSlide, targetSlide);
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

        public static readonly SyncFunction SyncVisualAgendaEndSlide = (refSlide, sections, currentSection, targetSlide) =>
        {
            DeleteVisualAgendaImageShapes(targetSlide);
            SyncSingleAgendaGeneral(refSlide, targetSlide);
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


        private static Shape FindShapeCorrespondingToSection(PowerPointSlide inSlide, int sectionIndex)
        {
            return inSlide.GetShape(AgendaShape.MeetsConditions(shape => shape.ShapePurpose == ShapePurpose.VisualAgendaImage &&
                                                                        sectionIndex == shape.Section.Index));
        }


        private static void GenerateVisualAgendaSlideZoomIn(PowerPointSlide slide, Shape zoomInShape)
        {
            // add drill down effect and clean up current slide by deleting drill down
            // shape and recover original slide shape visibility
            PowerPointDrillDownSlide addedSlide;
            AutoZoom.AddDrillDownAnimation(zoomInShape, slide, out addedSlide, includeAckSlide: false, deletePreviouslyAdded: false);
            slide.GetShapesWithRule(new Regex("PPTZoomIn"))[0].Delete();
            AgendaSection section = AgendaSlide.Decode(slide).Section;
            AgendaSlide.SetSlideName(addedSlide, Type.Visual, SlidePurpose.ZoomIn, section);
            zoomInShape.Visible = MsoTriState.msoTrue;
        }

        private static void GenerateVisualAgendaSlideZoomOut(PowerPointSlide slide, Shape zoomOutShape, bool finalZoomOut = false)
        {
            // add step back effect  and clean up current slide by deleting step back
            // shape and recover original slide shape visibility
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

        private static void SyncSingleAgendaGeneral(PowerPointSlide refSlide, PowerPointSlide candidate)
        {
            if (refSlide == null || candidate == null ||
                refSlide == candidate)
            {
                return;
            }

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
            var sameShapes = refSlide.Shapes.Cast<Shape>()
                                            .Where(shape => !PowerPointSlide.IsIndicator(shape) &&
                                                            !PowerPointSlide.IsTemplateSlideMarker(shape) &&
                                                            candidate.HasShapeWithSameName(shape.Name));

            foreach (var refShape in sameShapes)
            {
                var candidateShape = candidate.GetShapeWithName(refShape.Name)[0];

                Graphics.SyncShape(refShape, candidateShape);
            }
        }

        #endregion
    }
}
