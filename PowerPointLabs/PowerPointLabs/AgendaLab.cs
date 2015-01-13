using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPointLabs.Models;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs
{
    internal static class AgendaLab
    {
        private const string PptLabsAgendaSlideNameFormat = "PptLabs{0}Agenda{1}Slide {2}";
        private const string PptLabsAgendaTitleShapeName = "PptLabsAgendaTitle";
        private const string PptLabsAgendaContentShapeName = "PptLabsAgendaContent";
        private const string PptLabsAgendaSlideTypeSearchPattern = @"PptLabs(\w+)Agenda(?:Start|End)Slide";
        private const string PptLabsAgendaSectionName = "PptLabsAgendaSection";

        private const float VisualAgendaItemMargin = 0.05f;

        private static readonly Regex AgendaSlideSearchPattern = new Regex(PptLabsAgendaSlideTypeSearchPattern);

        private static readonly string SlideCapturePath = Path.Combine(Path.GetTempPath(), "PowerPointLabs");

        private static string _agendaText;

        # region Enum
        public enum AgendaType
        {
            Beam,
            Bullet,
            Visual
        };
        # endregion

        # region Properties
        public static bool HasAgenda
        {
            get
            {
                return PowerPointPresentation.Current.Slides.Any(slide => AgendaSlideSearchPattern.IsMatch(slide.Name));
            }
        }
        # endregion

        # region API
        public static void GenerateAgenda(AgendaType type)
        {
            // agenda exists in current presentation
            if (PowerPointPresentation.Current.Slides.Any(slide => AgendaSlideSearchPattern.IsMatch(slide.Name)))
            {
                var confirm = MessageBox.Show(TextCollection.AgendaLabAgendaExistError,
                                              TextCollection.AgendaLabAgendaExistErrorCaption,
                                              MessageBoxButtons.OKCancel);

                if (confirm == DialogResult.OK)
                {
                    RemoveAgenda();
                }
                else
                {
                    return;
                }
            }

            // validate section information
            var sections = PowerPointPresentation.Current.Sections.Skip(1).ToList();

            if (sections.Count == 0)
            {
                MessageBox.Show(TextCollection.AgendaLabNoSectionError);
                return;
            }

            if (sections.Count == 1)
            {
                MessageBox.Show(TextCollection.AgendaLabNoSectionError);
                return;
            }

            switch (type)
            {
                case AgendaType.Beam:
                    GenerateBeamAgenda();
                    break;
                case AgendaType.Bullet:
                    GenerateBulletAgenda(sections);
                    break;
                case AgendaType.Visual:
                    GenerateVisualAgenda(sections);
                    break;
            }
        }

        public static void RemoveAgenda()
        {
            var sectionProperties = PowerPointPresentation.Current.Presentation.SectionProperties;
            var section = PowerPointPresentation.Current.Sections;

            for (var i = section.Count; i >= 1; i --)
            {
                if (section[i - 1] == PptLabsAgendaSectionName)
                {
                    sectionProperties.Delete(i, true);
                }
            }

            var slides = PowerPointPresentation.Current.Slides;

            foreach (var slide in slides.Where(slide => AgendaSlideSearchPattern.IsMatch(slide.Name)))
            {
                slide.Delete();
            }
        }

        public static void SyncrhonizeAgenda()
        {
            // find the agenda for the first section as reference
            PowerPointSlide startRef;
            PowerPointSlide endRef;

            var type = FindSyncReference(out startRef, out endRef);

            SyncAgendaGeneral(startRef, endRef);

            switch (type)
            {
                case AgendaType.Beam:
                    break;
                case AgendaType.Bullet:
                    SyncAgendaBulletType(startRef, endRef);
                    break;
                case AgendaType.Visual:
                    break;
            }
        }
        # endregion

        # region Helper Functions
        private static void AddAgendaSlideBulletType(string section, bool isEnd)
        {
            var sectionIndex = FindSectionIndex(section);
            var sectionEndIndex = FindSectionEnd(section);

            var slide =
                PowerPointSlide.FromSlideFactory(PowerPointPresentation.Current
                                                                       .Presentation
                                                                       .Slides
                                                                       .Add(isEnd ? sectionEndIndex + 1 : 1,
                                                                            PpSlideLayout.ppLayoutText));

            slide.Name = string.Format(PptLabsAgendaSlideNameFormat, AgendaType.Bullet,
                                          isEnd ? "End" : "Start", section);
            slide.Transition.EntryEffect = PpEntryEffect.ppEffectFadeSmoothly;
            slide.Transition.Duration = 0.25f;

            slide.Shapes.Placeholders[1].Name = PptLabsAgendaTitleShapeName;
            slide.Shapes.Placeholders[2].Name = PptLabsAgendaContentShapeName;

            if (!isEnd)
            {
                slide.GetNativeSlide().MoveToSectionStart(sectionIndex);
            }

            // set title
            slide.Shapes.Placeholders[1].TextFrame.TextRange.Text = "Agenda";

            // set agenda content
            var contentPlaceHolder = slide.Shapes.Placeholders[2];
            var textRange = contentPlaceHolder.TextFrame.TextRange;

            // since we skip the default section, relative section index should be substracted by 1
            var relativeSectionIndex = sectionIndex - 1;

            textRange.Text = _agendaText;

            for (var i = 1; i < relativeSectionIndex; i++)
            {
                textRange.Paragraphs(i).Font.Color.RGB = PowerPointLabsGlobals.CreateRGB(Color.Gray);
            }

            textRange.Paragraphs(relativeSectionIndex).Font.Color.RGB =
                PowerPointLabsGlobals.CreateRGB(isEnd ? Color.Gray : Color.Red);
        }

        private static void AddAgendaSlideVisualType(List<string> sections)
        {
            // TODO: integrate these 2 parts together!!!

            // add a new section before the first section and rename to PptLabsAgendaSectionName
            var index = FindSectionStart(sections[0]);
            var currentPresentation = PowerPointPresentation.Current.Presentation;
            var sectionProperties = currentPresentation.SectionProperties;
            var slide = PowerPointSlide.FromSlideFactory(currentPresentation.Slides
                                                                            .Add(index - 1,
                                                                                 PpSlideLayout.ppLayoutTitleOnly));

            sectionProperties.AddBeforeSlide(index - 1, PptLabsAgendaSectionName);

            // generate slide shapes in the canvas area
            PrepareVisualAgendaSlideShapes(slide, sections);

            // get the shape that represent current slide
            var slideShape = slide.GetShapeWithName(sections[0])[0];

            // generate drill down slide, and clean up current slide by deleting drill down
            // shape and recover original slide shape visibility
            AutoZoom.AddDrillDownAnimation(slideShape, slide);
            slide.GetShapesWithRule(new Regex("PPTLabsZoomIn"))[0].Delete();
            slideShape.Visible = MsoTriState.msoTrue;

            // copy current slide for the next agenda section
            slide.Copy();

            // generate agenda for the rest of the sections
            foreach (var section in sections.Skip(1))
            {
                // add a new section before the first section and rename to PptLabsAgendaSectionName
                index = FindSectionStart(section);
                slide = PowerPointSlide.FromSlideFactory(currentPresentation.Slides.Paste(index - 1)[1]);

                sectionProperties.AddBeforeSlide(index - 1, PptLabsAgendaSectionName);

                // get the shape that represent current slide
                slideShape = slide.GetShapeWithName(section)[0];

                // add step back effect  and clean up current slide by deleting step back
                // shape and recover original slide shape visibility
                AutoZoom.AddStepBackAnimation(slideShape, slide);
                slide.GetShapesWithRule(new Regex("PPTLabsZoomIn"))[0].Delete();
                slideShape.Visible = MsoTriState.msoTrue;

                // add drill down effect and clean up current slide by deleting drill down
                // shape and recover original slide shape visibility
                AutoZoom.AddDrillDownAnimation(slideShape, slide);
                slide.GetShapesWithRule(new Regex("PPTLabsZoomIn"))[0].Delete();
                slideShape.Visible = MsoTriState.msoTrue;

                slide.Copy();
            }
        }

        private static int FindSectionEnd(string section)
        {
            var sectionIndex = FindSectionIndex(section);

            return FindSectionEnd(sectionIndex);
        }

        private static int FindSectionEnd(int sectionIndex)
        {
            // here the sectionIndex is 1-based!
            var sectionProperties = PowerPointPresentation.Current.Presentation.SectionProperties;

            return sectionProperties.FirstSlide(sectionIndex) + sectionProperties.SlidesCount(sectionIndex) - 1;
        }

        private static int FindSectionStart(string section)
        {
            var sectionIndex = FindSectionIndex(section);

            return FindSectionStart(sectionIndex);
        }

        private static int FindSectionStart(int sectionIndex)
        {
            // here the sectionIndex is 1-based!
            var sectionProperties = PowerPointPresentation.Current.Presentation.SectionProperties;

            return sectionProperties.FirstSlide(sectionIndex);
        }

        private static int FindSectionIndex(string section)
        {
            // here the return value is 1-based!
            return PowerPointPresentation.Current.Sections.FindIndex(name => name == section) + 1;
        }

        private static AgendaType FindSyncReference(out PowerPointSlide startRef, out PowerPointSlide endRef)
        {
            // the first meaningful section is the second section
            startRef = PowerPointPresentation.Current.Slides[FindSectionStart(2) - 1];
            endRef = PowerPointPresentation.Current.Slides[FindSectionEnd(2) - 1];

            var type = AgendaSlideSearchPattern.Match(startRef.Name).Groups[1].Value;

            return (AgendaType)Enum.Parse(typeof(AgendaType), type);
        }

        private static void GenerateBeamAgenda()
        {
            throw new NotImplementedException();
        }

        private static void GenerateBulletAgenda(List<string> sections)
        {
            // need to use '\r' as paragraph indicator, not '\n'!
            // must end with '\r' to make the last line a paragraph!
            _agendaText = sections.Aggregate((current, next) => current + "\r" + next) + "\r";

            foreach (var section in sections)
            {
                AddAgendaSlideBulletType(section, false);
                AddAgendaSlideBulletType(section, true);
            }
        }

        private static void GenerateVisualAgenda(List<string> sections)
        {
            PrepareVisualAgendaSlideCapture(sections);

            AddAgendaSlideVisualType(sections);
        }

        private static void PrepareVisualAgendaSlideCapture(IEnumerable<string> sections)
        {
            var slides = PowerPointPresentation.Current.Slides;

            foreach (var section in sections)
            {
                var sectionStartSlide = slides[FindSectionStart(section) - 1];
                var sectionEndSlide = slides[FindSectionEnd(section) - 1];
                var animatedEndSlide = sectionEndSlide.Duplicate();

                foreach (var shape in animatedEndSlide.Shapes.Cast<Shape>().Where(animatedEndSlide.HasExitAnimation))
                {
                    shape.Delete();
                }

                animatedEndSlide.MoveMotionAnimation();

                var sectionStartName = string.Format("{0} Start.png", section);
                var sectionEndName = string.Format("{0} End.png", section);

                Utils.Graphics.ExportSlide(sectionStartSlide, Path.Combine(SlideCapturePath, sectionStartName));
                Utils.Graphics.ExportSlide(animatedEndSlide, Path.Combine(SlideCapturePath, sectionEndName));

                animatedEndSlide.Delete();
            }
        }

        private static void PrepareVisualAgendaSlideShapes(PowerPointSlide slide, List<string> sections)
        {
            var titleBar = slide.Shapes.Placeholders[1];

            titleBar.Name = PptLabsAgendaTitleShapeName;
            titleBar.TextFrame.TextRange.Text = "Agenda";

            var slideWidth = PowerPointPresentation.Current.SlideWidth;
            var slideHeight = PowerPointPresentation.Current.SlideHeight;
            var aspectRatio = slideWidth / slideHeight;
            var epsilon = slideHeight * 0.02f;

            var canvasLeft = titleBar.Left;
            var canvasTop = titleBar.Top + titleBar.Height + epsilon;
            var canvasWidth = titleBar.Width;
            var canvasHeight = canvasWidth / aspectRatio;

            var itemCount = sections.Count;
            var itemCanvasWidth = canvasWidth / itemCount;
            var itemWidth = itemCanvasWidth * (1 - 2 * VisualAgendaItemMargin);
            var itemHeight = itemWidth / aspectRatio;
            var itemTop = canvasTop + (canvasHeight - itemHeight) / 2;
            
            for (var i = 0; i < itemCount; i ++)
            {
                var itemLeft = canvasLeft + i*itemCanvasWidth + itemCanvasWidth * VisualAgendaItemMargin;

                var shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,
                                                  itemLeft, itemTop,
                                                  itemWidth, itemHeight);

                shape.Name = sections[i];
                shape.Line.Visible = MsoTriState.msoFalse;

                var slideCaptureName = string.Format("{0} Start.png", shape.Name);
                shape.Fill.UserPicture(Path.Combine(SlideCapturePath, slideCaptureName));
            }
        }

        private static void SyncAgendaBulletType(PowerPointSlide startRef, PowerPointSlide endRef)
        {
            // in this step, we should sync:
            // 1. Content position;
            // 2. Content format.
            var sectionCount = PowerPointPresentation.Current.Sections.Count;

            // Section 1: default section, skip
            // Section 2: use as reference
            // Section 3 - end: need to be sync
            for (var i = 3; i <= sectionCount; i++)
            {
                var startAgenda = PowerPointPresentation.Current.Slides[FindSectionStart(i) - 1];
                var endAgenda = PowerPointPresentation.Current.Slides[FindSectionEnd(i) - 1];

                startAgenda.Layout = startRef.Layout;
                endAgenda.Layout = endRef.Layout;

                SyncSingleAgendaBullet(startRef, startAgenda);
                SyncSingleAgendaBullet(endRef, endAgenda);
            }
        }

        private static void SyncAgendaGeneral(PowerPointSlide startRef, PowerPointSlide endRef)
        {
            // in this step, we should sync:
            // 1. Layout
            // 2. Design;
            // 3. Transition;
            // 4. Shapes;
            // 5. Title Position
            // 6. TitleText
            var sectionCount = PowerPointPresentation.Current.Sections.Count;

            // Section 1: default section, skip
            // Section 2: use as reference
            // Section 3 - end: need to be sync
            for (var i = 3; i <= sectionCount; i++)
            {
                var startAgenda = PowerPointPresentation.Current.Slides[FindSectionStart(i) - 1];
                var endAgenda = PowerPointPresentation.Current.Slides[FindSectionEnd(i) - 1];

                startAgenda.Layout = startRef.Layout;
                endAgenda.Layout = endRef.Layout;

                SyncSingleAgendaGeneral(startRef, startAgenda);
                SyncSingleAgendaGeneral(endRef, endAgenda);
            }
        }

        private static void SyncSingleAgendaBullet(PowerPointSlide refSlide, PowerPointSlide candidate)
        {
            var refContentShape = refSlide.GetShapeWithName(PptLabsAgendaContentShapeName)[0];
            var candidateContentShape = candidate.GetShapeWithName(PptLabsAgendaContentShapeName)[0];

            candidateContentShape.Left = refContentShape.Left;
            candidateContentShape.Width = refContentShape.Width;

            var refParagraphCount = refContentShape.TextFrame2.TextRange.Paragraphs.Count;
            var candidateParagraphCount = candidateContentShape.TextFrame2.TextRange.Paragraphs.Count;
            var refTextRange = refContentShape.TextFrame.TextRange;
            var candidateTextRange = candidateContentShape.TextFrame.TextRange;

            if (refParagraphCount == candidateParagraphCount)
            {
                for (var i = 1; i <= refParagraphCount; i++)
                {
                    var refParagraph = refTextRange.Paragraphs(i);
                    var candidateParagraph = candidateTextRange.Paragraphs(i);
                    var candidateColor = candidateParagraph.Font.Color.RGB;

                    refParagraph.Copy();

                    var newCandidateRange = candidateParagraph.PasteSpecial();

                    newCandidateRange.Font.Color.RGB = candidateColor;
                }
            }
        }

        private static void SyncSingleAgendaGeneral(PowerPointSlide refSlide, PowerPointSlide candidate)
        {
            candidate.Design = refSlide.Design;
            candidate.Transition = refSlide.Transition;

            // syncronize extra shapes
            var extraShapes = refSlide.Shapes.Cast<Shape>()
                                             .Where(shape => shape.Name != PptLabsAgendaTitleShapeName &&
                                                             shape.Name != PptLabsAgendaContentShapeName &&
                                                             !candidate.HasShapeWithSameName(shape.Name))
                                             .Select(shape => shape.Name)
                                             .ToArray();

            try
            {
                if (extraShapes.Length != 0)
                {
                    refSlide.Shapes.Range(extraShapes).Copy();
                    candidate.Shapes.Paste();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }

            // syncronize title textbox
            var refTitleShape = refSlide.GetShapeWithName(PptLabsAgendaTitleShapeName)[0];
            var candidateTitleShape = candidate.GetShapeWithName(PptLabsAgendaTitleShapeName)[0];

            if (refTitleShape == null && candidateTitleShape != null)
            {
                candidateTitleShape.Delete();
            } else
            if (refTitleShape != null && candidateTitleShape == null)
            {
                refTitleShape.Copy();
                candidate.Shapes.Paste();
            } else
            if (refTitleShape != null)
            {
                candidateTitleShape.Left = refTitleShape.Left;
                candidateTitleShape.Width = refTitleShape.Width;
                candidateTitleShape.TextFrame.TextRange.Text = refTitleShape.TextFrame.TextRange.Text;

                refTitleShape.PickUp();
                candidateTitleShape.Apply();
            }
        }
        # endregion
    }
}
