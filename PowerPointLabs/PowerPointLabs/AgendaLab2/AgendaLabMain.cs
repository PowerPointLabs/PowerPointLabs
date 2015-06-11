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
    /// <summary>
    /// The sections should not change during generation / syncing.
    /// </summary>
    internal static partial class AgendaLabMain
    {
        private static LoadingDialog _loadDialog = new LoadingDialog();

        #region Bullet Formats
        private struct BulletFormats
        {
            public readonly TextRange2 Visited;
            public readonly TextRange2 Highlighted;
            public readonly TextRange2 Unvisited;

            private BulletFormats(TextRange2 visited, TextRange2 highlighted, TextRange2 unvisited)
            {
                Visited = visited;
                Highlighted = highlighted;
                Unvisited = unvisited;
            }

            /// <summary>
            /// Assumes the number of paragraphs >= 3.
            /// The check should have been done before this function is called.
            /// </summary>
            public static BulletFormats ExtractFormats(Shape contentShape)
            {
                var paragraphs = contentShape.TextFrame2.TextRange.Paragraphs.Cast<TextRange2>().ToList();
                return new BulletFormats(paragraphs[0],
                                        paragraphs[1],
                                        paragraphs[2]);
            }
        }
        #endregion

        #region API
        public static void GenerateAgenda(Type type)
        {
            bool dialogOpen = false;
            var curWindow = Globals.ThisAddIn.Application.ActiveWindow;
            var oldViewType = curWindow.ViewType;

            try
            {
                var slideTracker = new SlideSelectionTracker(SelectedSlides, CurrentSlide);

                if (AgendaPresent())
                {
                    var confirm = MessageBox.Show(TextCollection.AgendaLabAgendaExistError,
                                                  TextCollection.AgendaLabAgendaExistErrorCaption,
                                                  MessageBoxButtons.OKCancel);
                    if (confirm != DialogResult.OK) return;

                    RemoveAllAgendaItems(slideTracker);
                }

                if (!ValidSections()) return;

                slideTracker.DeleteAcknowledgementSlideAndTrack();

                dialogOpen = DisplayLoadingDialog(TextCollection.AgendaLabLoadingDialogTitle,
                                                    TextCollection.AgendaLabLoadingDialogContent);
                curWindow.ViewType = PpViewType.ppViewNormal;

                switch (type)
                {
                    case Type.Beam:
                        CreateBeamAgenda(slideTracker.SelectedSlides);
                        break;
                    case Type.Bullet:
                        CreateBulletAgenda(slideTracker);
                        break;
                    case Type.Visual:
                        CreateVisualAgenda(slideTracker);
                        break;
                }

                PowerPointPresentation.Current.AddAckSlide();
                SelectOriginalSlide(slideTracker.UserCurrentSlide, PowerPointPresentation.Current.FirstSlide);
            }
            finally
            {
                if (dialogOpen)
                {
                    DisposeLoadingDialog();
                }
                curWindow.ViewType = oldViewType;
            }
        }

        public static void RemoveAgenda()
        {
            var curWindow = Globals.ThisAddIn.Application.ActiveWindow;
            var oldViewType = curWindow.ViewType;

            try
            {
                var slideTracker = new SlideSelectionTracker(SelectedSlides, CurrentSlide);

                if (!AgendaPresent())
                {
                    MessageBox.Show(TextCollection.AgendaLabNoAgendaError);
                    return;
                }
                curWindow.ViewType = PpViewType.ppViewNormal;

                RemoveAllAgendaItems(slideTracker);

                SelectOriginalSlide(slideTracker.UserCurrentSlide, PowerPointPresentation.Current.FirstSlide);
            }
            finally
            {
                curWindow.ViewType = oldViewType;
            }
        }

        public static void SynchroniseAgenda()
        {
            bool dialogOpen = false;
            var curWindow = Globals.ThisAddIn.Application.ActiveWindow;
            var oldViewType = curWindow.ViewType;

            try
            {
                var slideTracker = new SlideSelectionTracker(SelectedSlides, CurrentSlide);
                var refSlide = FindReferenceSlide();
                var type = GetReferenceSlideType();

                if (!AgendaPresent())
                {
                    MessageBox.Show(TextCollection.AgendaLabNoAgendaError);
                    return;
                }
                if (refSlide == null)
                {
                    MessageBox.Show(TextCollection.AgendaLabNoReferenceSlideError);
                    return;
                }
                if (InvalidReferenceSlide(type, refSlide))
                {
                    MessageBox.Show(TextCollection.AgendaLabInvalidReferenceSlideError);
                    return;
                }
                if (!ValidSections()) return;

                slideTracker.DeleteAcknowledgementSlideAndTrack();
                dialogOpen = DisplayLoadingDialog(TextCollection.AgendaLabSynchronizingDialogTitle,
                                                    TextCollection.AgendaLabSynchronizingDialogContent);
                curWindow.ViewType = PpViewType.ppViewNormal;

                BringToFront(refSlide);

                switch (type)
                {
                    case Type.Beam:

                        break;
                    case Type.Bullet:
                        SyncBulletAgenda(slideTracker, refSlide);
                        break;
                    case Type.Visual:
                        RegenerateReferenceSlideImages(refSlide);
                        SyncVisualAgenda(slideTracker, refSlide);
                        break;
                }

                PowerPointPresentation.Current.AddAckSlide();
                SelectOriginalSlide(slideTracker.UserCurrentSlide, PowerPointPresentation.Current.FirstSlide);
            }
            finally
            {
                if (dialogOpen)
                {
                    DisposeLoadingDialog();
                }
                curWindow.ViewType = oldViewType;
            }
        }

        private static bool InvalidReferenceSlide(Type type, PowerPointSlide refSlide)
        {
            switch (type)
            {
                case Type.Beam:
                    return InvalidBeamAgendaReferenceSlide(refSlide);
                case Type.Bullet:
                    return InvalidBulletAgendaReferenceSlide(refSlide);
                case Type.Visual:
                    return InvalidVisualAgendaReferenceSlide(refSlide);
            }
            return true;
        }

        #endregion


        #region FUNCTIONS

        /// <summary>
        /// Assumption: no reference slide exists
        /// </summary>
        private static void CreateBulletAgenda(SlideSelectionTracker slideTracker)
        {
            var refSlide = CreateBulletReferenceSlide();

            // here we invoke sync logic, since it's the same behavior as sync
            SyncBulletAgenda(slideTracker, refSlide);
        }


        /// <summary>
        /// Assumption: no reference slide exists
        /// </summary>
        private static void CreateVisualAgenda(SlideSelectionTracker slideTracker)
        {
            var refSlide = CreateVisualReferenceSlide();

            // here we invoke sync logic, since it's the same behavior as sync
            SyncVisualAgenda(slideTracker, refSlide);
        }


        private static void SelectOriginalSlide(PowerPointSlide originalSlide, PowerPointSlide fallbackToSlide)
        {
            if (originalSlide != null)
            {
                originalSlide.GetNativeSlide().Select();
                return;
            }
            if (fallbackToSlide != null)
            {
                fallbackToSlide.GetNativeSlide().Select();
            }
        }


        private static void RemoveAllAgendaItems(SlideSelectionTracker slideTracker = null)
        {
            if (slideTracker == null) slideTracker = SlideSelectionTracker.CreateInactiveTracker();

            PowerPointPresentation.Current.Slides.Where(AgendaSlide.IsAnyAgendaSlide)
                                                .ToList()
                                                .ForEach(slideTracker.DeleteSlideAndTrack);

            RemoveBeamAgendaFromSlides(PowerPointPresentation.Current.Slides);
        }

        private static void RemoveBeamAgendaFromSlides(IEnumerable<PowerPointSlide> candidates)
        {
            candidates = candidates.Where(AgendaSlide.IsNotReferenceslide);
            foreach (var candidate in candidates)
            {
                var beamShape = FindBeamShape(candidate);

                if (beamShape != null)
                {
                    beamShape.Delete();
                }
            }
        }

        private static void BringToFront(PowerPointSlide slide)
        {
            slide.MoveTo(1);
        }

        private static void SyncBulletAgenda(SlideSelectionTracker slideTracker, PowerPointSlide refSlide)
        {
            var sections = Sections;

            ScrambleSlideSectionNames();
            foreach (var currentSection in sections)
            {
                var template = new BulletAgendaTemplate();
                ConfigureTemplate(currentSection, template);

                var templateTable = RebuildSectionUsingTemplate(slideTracker, currentSection, template);
                SynchroniseAllSlides(template, templateTable, refSlide, sections, currentSection);
            }
        }


        private static void SyncVisualAgenda(SlideSelectionTracker slideTracker, PowerPointSlide refSlide)
        {
            var sections = Sections;

            DeleteAllZoomSlides(slideTracker);
            ScrambleSlideSectionNames();
            foreach (var currentSection in sections)
            {
                var template = new VisualAgendaTemplate();
                ConfigureTemplate(currentSection, template);

                var templateTable = RebuildSectionUsingTemplate(slideTracker, currentSection, template);
                SynchroniseAllSlides(template, templateTable, refSlide, sections, currentSection);
            }
        }

        private static Dictionary<int, Shape> GetShapeAssignment(PowerPointSlide inSlide, out List<Shape> unassignedShapes)
        {
            var shapes = inSlide.Shapes.Cast<Shape>();

            unassignedShapes = new List<Shape>();
            var shapeAssignment = new Dictionary<int, Shape>();

            foreach (var shape in shapes)
            {
                var agendaShape = AgendaShape.Decode(shape);
                if (agendaShape == null || agendaShape.ShapePurpose != ShapePurpose.VisualAgendaImage) continue;

                int index = agendaShape.Section.Index;
                if (shapeAssignment.ContainsKey(index))
                {
                    unassignedShapes.Add(shape);
                }
                else
                {
                    shapeAssignment.Add(index, shape);
                }
            }

            return shapeAssignment;
        }

        private static void RegenerateReferenceSlideImages(PowerPointSlide refSlide)
        {
            List<Shape> markedForDeletion;
            var shapeAssignment = GetShapeAssignment(refSlide, out markedForDeletion);

            var sections = GetAllButFirstSection();
            var assignedOldIndexes = new HashSet<int>();
            var unassignedNewSections = new List<AgendaSection>();


            float existingImageWidth = -1;
            float existingImageHeight = -1;

            foreach (var section in sections)
            {
                int oldIndex = IdentifyOldSectionIndex(section);
                if (oldIndex == -1 || assignedOldIndexes.Contains(oldIndex))
                {
                    unassignedNewSections.Add(section);
                    continue;
                }
                Shape imageShape;
                bool canFindShape = shapeAssignment.TryGetValue(oldIndex, out imageShape);
                if (!canFindShape)
                {
                    unassignedNewSections.Add(section);
                    continue;
                }

                existingImageWidth = imageShape.Width;
                existingImageHeight = imageShape.Height;

                UpdateSectionImage(refSlide, section, imageShape);
                assignedOldIndexes.Add(oldIndex);
                
            }

            markedForDeletion.AddRange(from entry in shapeAssignment where !assignedOldIndexes.Contains(entry.Key) select entry.Value);

            var newSectionImages = 
                unassignedNewSections.Select(section => CreateSectionImage(refSlide, section))
                .ToList();
            PositionNewImageShapes(newSectionImages, existingImageWidth, existingImageHeight);

            markedForDeletion.ForEach(shape => shape.Delete());
        }

        private static int IdentifyOldSectionIndex(AgendaSection section)
        {
            var sectionSlides = GetSectionSlides(section);
            foreach (var slide in sectionSlides)
            {
                var agendaSlide = AgendaSlide.Decode(slide);
                if (agendaSlide != null)
                {
                    return agendaSlide.Section.Index;
                }
            }
            return -1;
        }

        private static void DeleteAllZoomSlides(SlideSelectionTracker slideTracker)
        {
            PowerPointPresentation.Current.Slides
                                        .Where(AgendaSlide.MeetsConditions(slide => slide.SlidePurpose == SlidePurpose.ZoomIn ||
                                                                                    slide.SlidePurpose == SlidePurpose.ZoomOut ||
                                                                                    slide.SlidePurpose == SlidePurpose.FinalZoomOut))
                                        .ToList()
                                        .ForEach(slideTracker.DeleteSlideAndTrack);
        }

        /// <summary>
        /// Scrambles the slide section names to avoid duplicate names later on, which can crash powerpoint.
        /// Use this just before reassigning the slide section names! Don't keep the slide names this way!
        /// </summary>
        private static void ScrambleSlideSectionNames()
        {
            var slides = PowerPointPresentation.Current.Slides;
            slides.Where(slide => AgendaSlide.IsAnyAgendaSlide(slide) && AgendaSlide.IsNotReferenceslide(slide))
                    .ToList()
                    .ForEach(AgendaSlide.AssignUniqueSectionName);
        }

        private static void SyncBeamAgenda()
        {
            SyncBeamAgendaSlides();
        }

        private static void SyncBeamAgendaSlides()
        {
            throw new NotImplementedException();
        }

        private static bool DisplayLoadingDialog(string title, string content)
        {
            _loadDialog = new LoadingDialog(title, content);
            _loadDialog.Show();
            _loadDialog.Refresh();
            return true;
        }

        private static void DisposeLoadingDialog()
        {
            _loadDialog.Dispose();
        }


        private static bool InvalidBulletAgendaReferenceSlide(PowerPointSlide refSlide)
        {
            var contentHolder = refSlide.GetShape(AgendaShape.WithPurpose(ShapePurpose.ContentShape));
            return (contentHolder == null || contentHolder.TextFrame2.TextRange.Paragraphs.Count < 3);
        }

        private static bool InvalidBeamAgendaReferenceSlide(PowerPointSlide refSlide)
        {
            throw new NotImplementedException();
        }

        private static bool InvalidVisualAgendaReferenceSlide(PowerPointSlide refSlide)
        {
            return false;
        }


        #endregion

        #region Sync Functions

        public static SyncFunction SyncVisualAgendaSlide = (refSlide, sections, currentSection, targetSlide) =>
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

        public static SyncFunction SyncVisualAgendaEndSlide = (refSlide, sections, currentSection, targetSlide) =>
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
            AutoZoom.AddDrillDownAnimation(zoomInShape, slide, out addedSlide, includeAckSlide: false);
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
            AutoZoom.AddStepBackAnimation(zoomOutShape, slide, out addedSlide, includeAckSlide: false);
            slide.GetShapesWithRule(new Regex("PPTZoomOut"))[0].Delete();
            AgendaSection section = AgendaSlide.Decode(slide).Section;
            AgendaSlide.SetSlideName(addedSlide, Type.Visual, finalZoomOut ? SlidePurpose.FinalZoomOut : SlidePurpose.ZoomOut, section);
            zoomOutShape.Visible = MsoTriState.msoTrue;

            var index = slide.Index;

            // move the step back slide to the first slide of the section
            PowerPointPresentation.Current.Presentation.Slides[index - 1].MoveTo(index);
            slide.MoveTo(index);
        }

        public static SyncFunction SyncBulletAgendaSlide = (refSlide, sections, currentSection, targetSlide) =>
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


        #region Conditions on current state

        private static bool ValidSections()
        {
            var sections = PowerPointPresentation.Current.Sections;

            if (sections.Count == 0)
            {
                MessageBox.Show(TextCollection.AgendaLabNoSectionError);
                return false;
            }

            if (sections.Count == 1)
            {
                MessageBox.Show(TextCollection.AgendaLabSingleSectionError);
                return false;
            }

            if (HasEmptySection())
            {
                MessageBox.Show(TextCollection.AgendaLabEmptySectionError);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Checks whether there is a section with no slides.
        /// Agenda slides are not counted.
        /// </summary>
        private static bool HasEmptySection()
        {
            var sections = Sections;
            foreach (var section in sections)
            {
                var sectionSlides = GetSectionSlides(section);
                if (sectionSlides.All(slide => AgendaSlide.IsAnyAgendaSlide(slide) || PowerPointAckSlide.IsAckSlide(slide)))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool IsReferenceSlidePresent()
        {
            return FindReferenceSlide() != null;
        }

        private static bool AgendaPresent()
        {
            return FindAllAgendaSlides().Count > 0 || FindSlidesWithBeam().Count > 0;
        }

        #endregion



        #region UTILITY
        /// <summary>
        /// If user has slides selected, return a list of the selected slides.
        /// If user is not currently selecting slides, returns an empty list.
        /// </summary>
        private static List<PowerPointSlide> SelectedSlides
        {
            get
            {
                if (PowerPointCurrentPresentationInfo.CurrentSelection.Type != PpSelectionType.ppSelectionSlides)
                    return new List<PowerPointSlide>();
                return PowerPointPresentation.Current.SelectedSlides;
            }
        }


        /// <summary>
        /// Return the slide the user is currently on.
        /// </summary>
        private static PowerPointSlide CurrentSlide
        {
            get
            {
                var slides = PowerPointPresentation.Current.SelectedSlides;
                if (slides.Count == 0) return null;
                return slides[0];
            }
        }

        private static void RemoveRemovedSlides(ref List<PowerPointSlide> selection, ref PowerPointSlide slideSelectedByUser, List<string> removedSlideNames)
        {
            selection = selection.Where(slide => !removedSlideNames.Contains(slide.Name)).ToList();
            slideSelectedByUser = removedSlideNames.Contains(slideSelectedByUser.Name) ? null : slideSelectedByUser;
        }


        private static PowerPointSlide FindReferenceSlide()
        {
            var slides = PowerPointPresentation.Current.Slides;
            return slides.FirstOrDefault(AgendaSlide.IsReferenceslide);
        }

        private static List<PowerPointSlide> FindAllAgendaSlides()
        {
            var slides = PowerPointPresentation.Current.Slides;
            return slides.Where(AgendaSlide.IsAnyAgendaSlide).ToList();
        }

        private static List<PowerPointSlide> FindSlidesWithBeam()
        {
            var slides = PowerPointPresentation.Current.Slides;
            return slides.Where(HasBeamShape).ToList();
        }

        private static Shape FindBeamShape(PowerPointSlide slide)
        {
            return slide.Shapes.Cast<Shape>().FirstOrDefault(AgendaShape.IsBeamShape);
        }

        private static bool HasBeamShape(PowerPointSlide slide)
        {
            return slide.Shapes.Cast<Shape>().Any(AgendaShape.IsBeamShape);
        }

        private static Type GetReferenceSlideType()
        {
            var referenceSlide = FindReferenceSlide();
            if (referenceSlide == null) return Type.None;
            return AgendaSlide.Decode(referenceSlide).AgendaType;
        }

        /// <summary>
        /// 1-indexed.
        /// </summary>
        private static List<AgendaSection> Sections
        {
            get
            {
                // TODO: This is a zip-with-index code. I can rephrase this few lines better (more functional)
                var sectionNames = PowerPointPresentation.Current.Sections;
                var sections = new List<AgendaSection>();
                for (var i = 0; i < sectionNames.Count; ++i)
                {
                    int index = i + 1;
                    sections.Add(new AgendaSection(sectionNames[i], index));
                }
                return sections;
            }
        }

        private static List<AgendaSection> GetAllButFirstSection()
        {
            var sections = Sections;
            if (sections.Count > 1)
            {
                sections.RemoveAt(0);
            }
            return sections;
        }


        /// <summary>
        /// Returns -1 when old section index is not found.
        /// </summary>
        private static int GetOldSectionIndex(AgendaSection section)
        {
            var sectionSlides = GetSectionSlides(section);
            foreach (var slide in sectionSlides)
            {
                var agendaSlide = AgendaSlide.Decode(slide);
                if (agendaSlide != null)
                {
                    return agendaSlide.Section.Index;
                }
            }
            return -1;
        }

        private static int NumberOfSections
        {
            get { return PowerPointPresentation.Current.Sections.Count; }
        }

        private static int FindSectionStart(AgendaSection section)
        {
            throw new NotImplementedException();
        }

        private static int FindSectionEnd(AgendaSection section)
        {
            throw new NotImplementedException();
        }

        private static AgendaSection FindSlideSection(PowerPointSlide slide)
        {
            throw new NotImplementedException();
        }

        private static PowerPointSlide FindSectionStartSlide(AgendaSection section)
        {
            return FindSectionStartSlide(section.Index);
        }

        private static PowerPointSlide FindSectionEndSlide(AgendaSection section)
        {
            return FindSectionEndSlide(section.Index);
        }

        private static PowerPointSlide FindSectionFirstNonAgendaSlide(int sectionIndex)
        {
            var slides = PowerPointPresentation.Current.Slides;
            int currentIndex = SectionFirstSlideIndex(sectionIndex) - 1;
            while (AgendaSlide.IsAnyAgendaSlide(slides[currentIndex]))
            {
                currentIndex++;
                if (currentIndex >= slides.Count)
                {
                    return null;
                }
            }
            return slides[currentIndex];
        }

        private static PowerPointSlide FindSectionLastNonAgendaSlide(int sectionIndex)
        {
            var slides = PowerPointPresentation.Current.Slides;
            int currentIndex = SectionLastSlideIndex(sectionIndex) - 1;
            while (AgendaSlide.IsAnyAgendaSlide(slides[currentIndex])) // TODO: ARGUMENTOUTOFRANGE
            {
                currentIndex--;
                if (currentIndex < 0)
                {
                    return null;
                }
            }
            return slides[currentIndex];
        }


        private static PowerPointSlide FindSectionStartSlide(int sectionIndex)
        {
            var slides = PowerPointPresentation.Current.Slides;
            int firstSlideIndex = SectionFirstSlideIndex(sectionIndex);
            return slides[firstSlideIndex - 1];
        }

        private static PowerPointSlide FindSectionEndSlide(int sectionIndex)
        {
            var slides = PowerPointPresentation.Current.Slides;
            int lastSlideIndex = SectionLastSlideIndex(sectionIndex);
            return slides[lastSlideIndex - 1];
        }

        /// <summary>
        /// 1-indexed.
        /// </summary>
        private static int SectionFirstSlideIndex(AgendaSection section)
        {
            return SectionFirstSlideIndex(section.Index);
        }

        /// <summary>
        /// 1-indexed
        /// </summary>
        private static int SectionLastSlideIndex(AgendaSection section)
        {
            return SectionLastSlideIndex(section.Index);
        }

        /// <summary>
        /// 1-indexed.
        /// </summary>
        private static int SectionFirstSlideIndex(int sectionIndex)
        {
            var sectionProperties = PowerPointPresentation.Current.SectionProperties;
            return sectionProperties.FirstSlide(sectionIndex);
        }

        /// <summary>
        /// 1-indexed
        /// </summary>
        private static int SectionLastSlideIndex(int sectionIndex)
        {
            var sectionProperties = PowerPointPresentation.Current.SectionProperties;
            int lastSlideIndex = PowerPointPresentation.Current.SlideCount;

            if (!IsLastSection(sectionIndex))
            {
                lastSlideIndex = sectionProperties.FirstSlide(sectionIndex + 1) - 1;
            }

            return lastSlideIndex;
        }

        /// <summary>
        /// 0-indexed.
        /// </summary>
        private static List<PowerPointSlide> GetSectionSlides(AgendaSection section, bool excludeReferenceSlide = false)
        {
            var slides = PowerPointPresentation.Current.Slides;

            int firstSlideIndex = SectionFirstSlideIndex(section);
            int lastSlideIndex = SectionLastSlideIndex(section);

            return slides.GetRange(firstSlideIndex - 1, lastSlideIndex - firstSlideIndex + 1);
        }

        private static bool IsLastSection(int sectionIndex)
        {
            return sectionIndex == NumberOfSections;
        }

        #endregion



        #region Agenda Creation

        /// <summary>
        /// Generates the beam agenda on the target slides. Skips over the Reference (Template) slide if included in targetSlides.
        /// Generates the Reference slide if it does not already exist.
        /// Leave the targetSlides field blank (=null) to generate the beam agenda over all slides (other than the first section).
        /// </summary>
        private static void CreateBeamAgenda(IEnumerable<PowerPointSlide> targetSlides = null)
        {
            var sections = Sections;

            List<PowerPointSlide> slides;
            if (targetSlides != null)
            {
                slides = targetSlides.Where(AgendaSlide.IsNotReferenceslide).ToList();
            }
            else
            {
                var firstSectionIndex = FindSectionStart(sections[0]);
                slides = PowerPointPresentation.Current.Slides
                    .Where(slide => slide.Index >= firstSectionIndex && AgendaSlide.IsNotReferenceslide(slide))
                    .ToList();
            }

            if (slides.Count < 1) return;

            var refSlide = FindReferenceSlide();
            bool generateNewReferenceSlide;
            if (AgendaSlide.IsReferenceslide(refSlide))
            {
                var beamShape = FindBeamShape(refSlide);
                if (beamShape != null)
                {
                    generateNewReferenceSlide = false;
                    beamShape.Copy();
                }
                else
                {
                    // reference slide doesn't have a beam shape. weird. so we delete and recreate.
                    generateNewReferenceSlide = true;
                    refSlide.Delete();
                }
            }
            else
            {
                // can't find a reference slide.
                generateNewReferenceSlide = true;
            }

            if (generateNewReferenceSlide)
            {
                // if we do not have legacy template, create a new refslide 
                refSlide = CreateBeamReferenceSlide();

                PrepareBeamAgendaShapes(sections, refSlide);
                AddAgendaSlideBeamType(sections[0], refSlide);
                refSlide.BringIndicatorToFront();
            }

            // The beam shape is now stored in the clipboard to be pasted on each of the slides.
            foreach (var slide in slides)
            {
                AddAgendaSlideBeamType(FindSlideSection(slide), slide);
            }
        }


        private static PowerPointSlide CreateBeamReferenceSlide()
        {
            var refSlide = PowerPointSlide.FromSlideFactory(PowerPointPresentation.Current
                                                            .Presentation
                                                            .Slides
                                                            .Add(1, PpSlideLayout.ppLayoutBlank));
            AgendaSlide.SetAsReferenceSlideName(refSlide, Type.Beam);
            refSlide.AddTemplateSlideMarker();
            refSlide.Hidden = true;

            return refSlide;
        }

        private static PowerPointSlide CreateBulletReferenceSlide()
        {
            var refSlide = PowerPointSlide.FromSlideFactory(PowerPointPresentation.Current
                                                            .Presentation
                                                            .Slides
                                                            .Add(1, PpSlideLayout.ppLayoutText));

            var titleShape = refSlide.Shapes.Placeholders[1];
            var contentShape = refSlide.Shapes.Placeholders[2];
            AgendaShape.SetShapeName(titleShape, ShapePurpose.TitleShape, AgendaSection.None);
            AgendaShape.SetShapeName(contentShape, ShapePurpose.ContentShape, AgendaSection.None);

            Graphics.SetText(titleShape, TextCollection.AgendaLabTitleContent);
            Graphics.SetText(contentShape, TextCollection.AgendaLabBulletVisitedContent,
                                            TextCollection.AgendaLabBulletHighlightedContent,
                                            TextCollection.AgendaLabBulletUnvisitedContent);

            var paragraphs = Graphics.GetParagraphs(contentShape);
            paragraphs[0].Font.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(Color.Gray);
            paragraphs[1].Font.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(Color.Red);
            paragraphs[2].Font.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(Color.Black);

            AgendaSlide.SetAsReferenceSlideName(refSlide, Type.Bullet);
            refSlide.AddTemplateSlideMarker();
            refSlide.Hidden = true;
            refSlide.DeleteIndicator();

            return refSlide;
        }


        private static PowerPointSlide CreateVisualReferenceSlide()
        {
            var refSlide = PowerPointSlide.FromSlideFactory(PowerPointPresentation.Current
                                                            .Presentation
                                                            .Slides
                                                            .Add(1, PpSlideLayout.ppLayoutTitleOnly));

            var titleBar = refSlide.Shapes.Placeholders[1];
            AgendaShape.SetShapeName(titleBar, ShapePurpose.TitleShape, AgendaSection.None);
            Graphics.SetText(titleBar, TextCollection.AgendaLabTitleContent);

            InsertVisualAgendaSectionImages(refSlide);

            AgendaSlide.SetAsReferenceSlideName(refSlide, Type.Visual);
            refSlide.AddTemplateSlideMarker();
            refSlide.Hidden = true;
            refSlide.DeleteIndicator();

            return refSlide;
        }

        /// <summary>
        /// Places the newly generated image shapes in some alignment that makes them easy to drag around.
        /// Resizes image shapes to match the sizes of the existing image shapes.
        /// If existingImageWidth <= 0 or existingImageHeight <= 0, it means there are no already existing image shapes.
        /// </summary> 
        private static void PositionNewImageShapes(List<Shape> shapes, float existingImageWidth, float existingImageHeight)
        {
            ArrangeInGrid(shapes);
            if (existingImageWidth <= 0 || existingImageHeight <= 0) return;
            
            foreach (var shape in shapes)
            {
                shape.Width = existingImageWidth;
                shape.Height = existingImageHeight;
            }
        }

        /// <summary>
        /// Inserts the section images into the reference slide in a nice square pattern and names them appropriately.
        /// </summary>
        private static void InsertVisualAgendaSectionImages(PowerPointSlide refSlide)
        {
            var sectionImages = CreateSectionImages(refSlide);
            ArrangeInGrid(sectionImages);
        }

        private static void ArrangeInGrid(List<Shape> sectionImages)
        {
            float slideWidth = PowerPointPresentation.Current.SlideWidth;
            float slideHeight = PowerPointPresentation.Current.SlideHeight;
            float aspectRatio = slideWidth/slideHeight;

            // These numbers can be tweaked.
            float panelFillRatio = 0.9f;
            float canvasTop = slideHeight*0.25f;
            float canvasBottom = slideHeight*0.85f;

            float canvasHeight = canvasBottom - canvasTop;
            float canvasWidth = aspectRatio*canvasHeight;
            float canvasLeft = (slideWidth - canvasWidth)/2;

            int columnCount = (int) Math.Ceiling(Math.Sqrt(sectionImages.Count));
            int rowCount = Common.CeilingDivide(sectionImages.Count, columnCount);
            float panelWidth = canvasWidth/columnCount;
            float panelHeight = panelWidth/aspectRatio;

            float pictureWidth = panelFillRatio*panelWidth;
            float pictureHeight = panelFillRatio*panelHeight;
            float pictureXOffset = canvasLeft + (panelWidth - pictureWidth)/2;
            float pictureYOffset = canvasTop + (canvasHeight - rowCount*panelHeight)/2 + (panelHeight - pictureHeight)/2;

            for (int i = 0; i < sectionImages.Count; ++i)
            {
                var sectionImage = sectionImages[i];
                int xPosition = i%columnCount;
                int yPosition = i/columnCount;

                sectionImage.Left = pictureXOffset + xPosition*panelWidth;
                sectionImage.Top = pictureYOffset + yPosition*panelHeight;
                sectionImage.Width = pictureWidth;
                sectionImage.Height = pictureHeight;
            }
        }

        private static List<Shape> CreateSectionImages(PowerPointSlide refSlide)
        {
            var sections = GetAllButFirstSection();
            var sectionImages = new List<Shape>();
            foreach (var section in sections)
            {
                var sectionImage = CreateSectionImage(refSlide, section);
                sectionImages.Add(sectionImage);
            }
            return sectionImages;
        }

        private static Shape CreateSectionImage(PowerPointSlide refSlide, AgendaSection section)
        {
            var sectionFirstSlide = FindSectionFirstNonAgendaSlide(section.Index);
            var shape = refSlide.InsertEntrySnapshotOfSlide(sectionFirstSlide);
            AgendaShape.SetShapeName(shape, ShapePurpose.VisualAgendaImage, section);
            return shape;
        }


        private static void UpdateSectionImage(PowerPointSlide refSlide, AgendaSection section, Shape imageShape)
        {
            var snapshotShape = CreateSectionImage(refSlide, section);
            Graphics.SyncShape(imageShape, snapshotShape, pickupShapeFormat: false, pickupTextContent: false, pickupTextFormat: false);
            imageShape.Delete();
        }


        private static void PrepareBeamAgendaShapes(List<AgendaSection> sections, PowerPointSlide refSlide)
        {
            throw new NotImplementedException();
        }

        private static void AddAgendaSlideBeamType(AgendaSection section, PowerPointSlide slide)
        {
            throw new NotImplementedException();
        }


        private static string CreateInDocHyperLink(PowerPointSlide slide)
        {
            throw new NotImplementedException();
        }
        #endregion

    }

}
