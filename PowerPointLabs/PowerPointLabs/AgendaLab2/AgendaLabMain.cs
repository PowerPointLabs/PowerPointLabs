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
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

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

        #region Beam Formats
        private struct BeamFormats
        {
            public readonly TextRange2 Highlighted;

            private BeamFormats(TextRange2 highlighted)
            {
                Highlighted = highlighted;
            }

            /// <summary>
            /// Assumes that beamTexts exists.
            /// </summary>
            public static BeamFormats ExtractFormats(Shape beamShape)
            {
                var groupItems = beamShape.GroupItems.Cast<Shape>();
                var beamTexts = groupItems.FirstOrDefault(AgendaShape.WithPurpose(ShapePurpose.BeamShapeHighlightedText));
                return new BeamFormats(beamTexts.TextFrame2.TextRange);
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
                        CreateBeamAgenda(slideTracker);
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


        #region Actions - Creation

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


        /// <summary>
        /// Assumption: no reference slide exists
        /// </summary>
        private static void CreateBeamAgenda(SlideSelectionTracker slideTracker)
        {
            var refSlide = CreateBeamReferenceSlide();

            // here we invoke sync logic, since it's the same behavior as sync
            var selectedSlides = slideTracker.SelectedSlides;
            if (selectedSlides.Count == 0)
            {
                selectedSlides = AllSlidesAfterFirstSection();
            }

            SyncBeamAgenda(selectedSlides, refSlide);
        }

        private static void SyncBeamAgenda(List<PowerPointSlide> selectedSlides, PowerPointSlide refSlide)
        {
            var sections = GetAllButFirstSection();
            var syncSlides = new List<PowerPointSlide>();

            // Generate beam agenda for all selected slides that do not currently have the beam agenda.
            if (selectedSlides != null)
            {
                var selectedSlidesWithoutBeam = selectedSlides.Where(slide => !HasBeamShape(slide));
                syncSlides.AddRange(selectedSlidesWithoutBeam);
            }
            // Synchronise agenda for all slides in the presentation that have the beam agenda.
            var refBeamShape = FindBeamShape(refSlide);
            var allSlidesWithBeam = PowerPointPresentation.Current.Slides
                                                                  .Where(slide => AgendaSlide.IsNotReferenceslide(slide) &&
                                                                                  FindBeamShape(slide) != null);
            syncSlides.AddRange(allSlidesWithBeam);

            foreach (var slide in syncSlides)
            {
                UpdateBeamAgenda(slide, refBeamShape);
            }
            
        }

        private static void UpdateBeamAgenda(PowerPointSlide slide, Shape refBeamShape)
        {
            RemoveBeamAgendaFromSlide(slide);
            refBeamShape.Copy();
            var beamShape = slide.Shapes.Paste();
            var section = GetSlideSection(slide);

            beamShape.GroupItems.Cast<Shape>()
                                .Where(AgendaShape.WithPurpose(ShapePurpose.BeamShapeHighlightedText))
                                .ToList()
                                .ForEach(shape => shape.Delete());

            if (section.Index == 1) return;

            var beamFormats = BeamFormats.ExtractFormats(refBeamShape);
            var currentSectionText = beamShape.GroupItems
                                            .Cast<Shape>()
                                            .Where(AgendaShape.MeetsConditions(shape => shape.ShapePurpose == ShapePurpose.BeamShapeText &&
                                                                                        shape.Section.Index == section.Index))
                                            .FirstOrDefault().TextFrame2.TextRange;

            Graphics.SyncTextRange(beamFormats.Highlighted, currentSectionText, pickupTextContent: false);
        }


        #endregion


        #region Reference Slide Creation - Beam

        private static PowerPointSlide CreateBeamReferenceSlide()
        {
            var refSlide = PowerPointSlide.FromSlideFactory(PowerPointPresentation.Current
                                                            .Presentation
                                                            .Slides
                                                            .Add(1, PpSlideLayout.ppLayoutBlank));

            CreateBeamAgendaShapes(refSlide);

            AgendaSlide.SetAsReferenceSlideName(refSlide, Type.Beam);
            refSlide.AddTemplateSlideMarker();
            refSlide.Hidden = true;

            return refSlide;
        }

        private static void CreateBeamAgendaShapes(PowerPointSlide refSlide, Direction beamDirection = Direction.Top)
        {
            var sections = GetAllButFirstSection();

            var lastLeft = 0.0f;
            var lastTop = 0.0f;
            var slideWidth = PowerPointPresentation.Current.SlideWidth;
            var slideHeight = PowerPointPresentation.Current.SlideHeight;

            var background = PrepareBeamAgendaBackground(refSlide);
            var widest = 0.0f;

            var textBoxes = new List<Shape>();
            foreach (var section in sections)
            {
                var textBox = PrepareBeamAgendaBeamItem(refSlide, lastLeft, lastTop, section);
                AdjustBeamItemHorizontal(ref lastLeft, ref lastTop, ref widest, 0, textBox, background);
                textBoxes.Add(textBox);
            }

            background.Height = 0;
            lastLeft = 0;
            lastTop = 0;
            var delta = Math.Max(widest, slideWidth / sections.Count);
            foreach (var textBox in textBoxes)
            {
                AdjustBeamItemHorizontal(ref lastLeft, ref lastTop, ref widest, delta, textBox, background);
            }

            var highlightedTextBox = CreateHighlightedTextBox(0, 10f, refSlide);

            var beamShapeItems = new List<Shape>();
            beamShapeItems.Add(background);
            beamShapeItems.Add(highlightedTextBox);
            beamShapeItems.AddRange(textBoxes);

            var group = refSlide.GroupShapes(beamShapeItems);
            AgendaShape.SetShapeName(group, ShapePurpose.BeamShapeMainGroup, AgendaSection.None);
        }

        private static Shape CreateHighlightedTextBox(float left, float top, PowerPointSlide slide)
        {
            var textBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                                                  left, top, 0, 0);

            AgendaShape.SetShapeName(textBox, ShapePurpose.BeamShapeHighlightedText, AgendaSection.None);
            textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            textBox.TextFrame.WordWrap = MsoTriState.msoFalse;
            textBox.TextFrame.TextRange.Text = TextCollection.AgendaLabBeamHighlightedText;
            textBox.TextFrame.TextRange.Font.Color.RGB = Graphics.ConvertColorToRgb(Color.Yellow);

            return textBox;
        }

        private static void AdjustBeamItemHorizontal(ref float lastLeft, ref float lastTop, ref float widest,
                                                     float delta, Shape item, Shape background)
        {
            if (lastLeft + delta > PowerPointPresentation.Current.SlideWidth)
            {
                lastLeft = 0;
                lastTop += item.Height;
            }

            item.Left = Math.Max(lastLeft, lastLeft + (delta - item.Width) / 2f);
            item.Top = lastTop;

            if (item.Width > widest)
            {
                widest = item.Width;
            }

            lastLeft += Math.Max(item.Width, delta);

            if (background.Height < lastTop + item.Height)
            {
                background.Height = lastTop + item.Height;
            }
        }


        private static Shape PrepareBeamAgendaBackground(PowerPointSlide slide)
        {
            var background = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 0, 0);
            background.Line.Visible = MsoTriState.msoFalse;
            background.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(Color.Black);
            background.Width = PowerPointPresentation.Current.SlideWidth;

            return background;
        }

        private static Shape PrepareBeamAgendaBeamItem(PowerPointSlide slide, float lastLeft, float lastTop, AgendaSection section)
        {
            var textBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                                                  lastLeft, lastTop, 0, 0);

            AgendaShape.SetShapeName(textBox, ShapePurpose.BeamShapeText, section);
            textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            textBox.TextFrame.WordWrap = MsoTriState.msoFalse;
            textBox.TextFrame.TextRange.Text = section.Name;
            textBox.TextFrame.TextRange.Font.Color.RGB = Graphics.ConvertColorToRgb(Color.White);

            return textBox;
        }

        #endregion


        #region Reference Slide Creation - Bullet

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

            return refSlide;
        }

        #endregion


        #region Reference Slide Creation - Visual

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

            return refSlide;
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
        
        #endregion


        #region Actions - Removal

        private static void RemoveAllAgendaItems(SlideSelectionTracker slideTracker = null)
        {
            if (slideTracker == null) slideTracker = SlideSelectionTracker.CreateInactiveTracker();

            PowerPointPresentation.Current.Slides.Where(AgendaSlide.IsAnyAgendaSlide)
                                                 .ToList()
                                                 .ForEach(slideTracker.DeleteSlideAndTrack);

            PowerPointPresentation.Current.Slides.ToList()
                                                 .ForEach(RemoveBeamAgendaFromSlide);
        }

        private static void RemoveBeamAgendaFromSlide(PowerPointSlide slide)
        {
            slide.Shapes.Cast<Shape>()
                        .Where(AgendaShape.WithPurpose(ShapePurpose.BeamShapeMainGroup))
                        .ToList()
                        .ForEach(shape => shape.Delete());
        }

        #endregion


        #region Actions - Synchronise

        private static void BringToFront(PowerPointSlide slide)
        {
            slide.MoveTo(1);
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

        private static void RegenerateReferenceSlideImages(PowerPointSlide refSlide)
        {
            List<Shape> markedForDeletion;
            var shapeAssignment = GetImageShapeAssignment(refSlide, out markedForDeletion);

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

        private static Dictionary<int, Shape> GetImageShapeAssignment(PowerPointSlide inSlide, out List<Shape> unassignedShapes)
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
        /// Identifies the previous section index of a section by looking at the generated agenda slides in the section.
        /// The identified section is the section index of the first generated agenda slide found.
        /// Returns -1 when old section index is not found.
        /// </summary>
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

        #endregion


        #region Actions - General

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


        private static bool InvalidBulletAgendaReferenceSlide(PowerPointSlide refSlide)
        {
            var contentHolder = refSlide.GetShape(AgendaShape.WithPurpose(ShapePurpose.ContentShape));
            return (contentHolder == null || contentHolder.TextFrame2.TextRange.Paragraphs.Count < 3);
        }

        private static bool InvalidBeamAgendaReferenceSlide(PowerPointSlide refSlide)
        {
            return !HasBeamShape(refSlide);
        }

        private static bool InvalidVisualAgendaReferenceSlide(PowerPointSlide refSlide)
        {
            return false;
        }

        #endregion


        private static string CreateInDocHyperLink(PowerPointSlide slide)
        {
            throw new NotImplementedException();
        }
    }

}
