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
using PowerPointLabs.Views;
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

        private static readonly string SlideCapturePath = Path.Combine(Path.GetTempPath(), "PowerPointLabs Temp");

        private static string _agendaText;

        private static Color _bulletDefaultColor = Color.Black;
        private static Color _bulletHighlightColor = Color.Red;
        private static Color _bulletDimColor = Color.Gray;

        # region Enum
        public enum AgendaType
        {
            None,
            Bullet,
            Beam,
            Visual
        };
        # endregion

        /*********************************************************
         * Note:
         * The implementation of CurrentAgendaType is O(n),
         * bear this in mind when design functions. E.g. 
         * FindSectionEndSlide is totally fine with only the first
         * argument and check CurrentAgendaType in the function.
         * But taking type as the second argument allows the 
         * developer to get CurrentAgendaType before the function.
         * This is useful when CurrentAgendaType is frequently used
         * in a certain region, which reduces the overhead.
         * 
         * See SyncrhonizeAgenda() for more usage details.
         *********************************************************/
        # region Properties
        public static AgendaType CurrentAgendaType
        {
            get
            {
                var agendaSlide = PowerPointPresentation.Current
                                                        .Slides
                                                        .FirstOrDefault(slide => AgendaSlideSearchPattern.
                                                                                 IsMatch(slide.Name));

                if (agendaSlide == null)
                {
                    return AgendaType.None;
                }

                var type = AgendaSlideSearchPattern.Match(agendaSlide.Name).Groups[1].Value;

                return (AgendaType)Enum.Parse(typeof(AgendaType), type);
            }
        }
        # endregion

        # region API
        public static void AgendaLabSettings()
        {
            var settingDialog = new BulletAgendaSettingsDialog(_bulletHighlightColor,
                                                               _bulletDimColor,
                                                               _bulletDefaultColor);

            settingDialog.SettingsHandler += UpdateColorScheme;
            settingDialog.ShowDialog();
        }

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
            if (CurrentAgendaType == AgendaType.None)
            {
                MessageBox.Show(TextCollection.AgendaLabNoAgendaError);
                return;
            }

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
            var type = CurrentAgendaType;

            if (type == AgendaType.None)
            {
                MessageBox.Show(TextCollection.AgendaLabNoAgendaError);
                return;
            }

            // find the agenda for the first section as reference
            var sections = PowerPointPresentation.Current.Sections;
            var refSlide = FindSectionStartSlide(sections[1], type);

            // Section 1: default section, skip
            // Section 2 start: use as reference
            // Section 2_end - end: need to be synced
            for (var i = 2; i <= sections.Count; i++)
            {
                var section = sections[i - 1];
                var startAgenda = i == 2 ? null : FindSectionStartSlide(section, type);
                var endAgenda = FindSectionEndSlide(section, type);

                SyncSingleAgendaGeneral(refSlide, startAgenda);
                SyncSingleAgendaGeneral(refSlide, endAgenda);

                switch (CurrentAgendaType)
                {
                    case AgendaType.Beam:
                        break;
                    case AgendaType.Bullet:
                        SyncSingleAgendaBullet(refSlide, startAgenda);
                        SyncSingleAgendaBullet(refSlide, endAgenda);
                        break;
                    case AgendaType.Visual:
                        SyncSingleAgendaVisual(refSlide, startAgenda);
                        SyncSingleAgendaVisual(refSlide, endAgenda);
                        break;
                }
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
            if (!isEnd)
            {
                slide.GetNativeSlide().MoveToSectionStart(sectionIndex);
            }

            slide.Name = string.Format(PptLabsAgendaSlideNameFormat, AgendaType.Bullet,
                                          isEnd ? "End" : "Start", section);
            slide.Transition.EntryEffect = PpEntryEffect.ppEffectFadeSmoothly;
            slide.Transition.Duration = 0.25f;

            slide.Shapes.Placeholders[1].Name = PptLabsAgendaTitleShapeName;
            slide.Shapes.Placeholders[2].Name = PptLabsAgendaContentShapeName;

            // set title
            slide.Shapes.Placeholders[1].TextFrame.TextRange.Text = "Agenda";

            // set agenda content
            var contentPlaceHolder = slide.Shapes.Placeholders[2];
            var textRange = contentPlaceHolder.TextFrame.TextRange;
            var focusColor = isEnd ? _bulletDimColor : _bulletHighlightColor;

            // since section index is 1-based, focus section index should be substracted by 1
            RecolorTextRange(textRange, sectionIndex - 1, focusColor);
        }

        private static void AddAgendaSlideVisualType(List<string> sections)
        {
            // TODO: integrate these 3 parts together!!!

            // add a new section before the first section and rename to PptLabsAgendaSectionName
            var index = FindSectionStart(sections[0]);
            var currentPresentation = PowerPointPresentation.Current.Presentation;
            var sectionProperties = currentPresentation.SectionProperties;
            var slide = PowerPointSlide.FromSlideFactory(currentPresentation.Slides
                                                                            .Add(index,
                                                                                 PpSlideLayout.ppLayoutTitleOnly));

            slide.Name = string.Format(PptLabsAgendaSlideNameFormat, AgendaType.Visual, "", sections[0]);
            sectionProperties.AddBeforeSlide(index, PptLabsAgendaSectionName);

            // generate slide shapes in the canvas area
            PrepareVisualAgendaSlideShapes(slide, sections);

            // get the shape that represent current slide
            var slideShape = slide.GetShapeWithName(sections[0])[0];

            // generate drill down slide, and clean up current slide by deleting drill down
            // shape and recover original slide shape visibility
            AutoZoom.AddDrillDownAnimation(slideShape, slide);
            slide.GetShapesWithRule(new Regex("PPTZoomIn"))[0].Delete();
            slideShape.Visible = MsoTriState.msoTrue;

            // copy current slide for the next agenda section
            slide.Copy();

            // generate agenda for the rest of the sections
            for (var i = 1; i <= sections.Count; i ++)
            {
                var sectionName = i < sections.Count ? sections[i] : sections[i - 1];
                var prevSectionName = sections[i - 1];

                // add a new section before the first section and rename to PptLabsAgendaSectionName
                index = i < sections.Count ? FindSectionStart(sectionName) : PowerPointPresentation.Current.SlideCount;
                slide = PowerPointSlide.FromSlideFactory(currentPresentation.Slides.Paste(index)[1]);

                slide.Name = string.Format(PptLabsAgendaSlideNameFormat, AgendaType.Visual, "", i < sections.Count ? sectionName : "EndOfAgenda");
                var newSectionIndex = sectionProperties.AddBeforeSlide(index, PptLabsAgendaSectionName);

                // get the shape that represent previous slide, change the fillup picture to
                // the end of slide picture
                var sectionEndName = string.Format("{0} End.png", prevSectionName);

                slideShape = slide.GetShapeWithName(prevSectionName)[0];
                slideShape.Fill.UserPicture(Path.Combine(SlideCapturePath, sectionEndName));

                // add step back effect  and clean up current slide by deleting step back
                // shape and recover original slide shape visibility
                AutoZoom.AddStepBackAnimation(slideShape, slide);
                slide.GetShapesWithRule(new Regex("PPTZoomOut"))[0].Delete();
                slideShape.Visible = MsoTriState.msoTrue;

                // move the step back slide to the section
                currentPresentation.Slides[index].MoveToSectionStart(newSectionIndex);

                if (i == sections.Count) break;

                // get the shape that represent current slide
                slideShape = slide.GetShapeWithName(sectionName)[0];

                // add drill down effect and clean up current slide by deleting drill down
                // shape and recover original slide shape visibility
                AutoZoom.AddDrillDownAnimation(slideShape, slide);
                slide.GetShapesWithRule(new Regex("PPTZoomIn"))[0].Delete();
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

        private static PowerPointSlide FindSectionEndSlide(string section, AgendaType type)
        {
            var slideName = string.Format(PptLabsAgendaSlideNameFormat, type, "End", section);

            return PowerPointPresentation.Current.Slides.FirstOrDefault(slide => slide.Name == slideName);
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

        private static PowerPointSlide FindSectionStartSlide(string section, AgendaType type)
        {
            var slideName = string.Format(PptLabsAgendaSlideNameFormat, type, "Start", section);

            return PowerPointPresentation.Current.Slides.FirstOrDefault(slide => slide.Name == slideName);
        }

        private static int FindSectionIndex(string section)
        {
            // here the return value is 1-based!
            return PowerPointPresentation.Current.Sections.FindIndex(name => name == section) + 1;
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
            if (!Directory.Exists(SlideCapturePath))
            {
                Directory.CreateDirectory(SlideCapturePath);
            }

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

        private static void RecolorTextRange(TextRange textRange, int focusIndex, Color focusColor)
        {
            textRange.Font.Color.RGB = PowerPointLabsGlobals.CreateRGB(_bulletDefaultColor);

            textRange.Text = _agendaText;

            for (var i = 1; i < focusIndex; i++)
            {
                textRange.Paragraphs(i).Font.Color.RGB = PowerPointLabsGlobals.CreateRGB(_bulletDimColor);
            }

            textRange.Paragraphs(focusIndex).Font.Color.RGB = PowerPointLabsGlobals.CreateRGB(focusColor);
        }

        private static void SyncShape(Shape refShape, Shape candidateShape,
                                      bool pickupFormat = true, bool pickupText = true)
        {
            candidateShape.Left = refShape.Left;
            candidateShape.Top = refShape.Top;
            candidateShape.Width = refShape.Width;
            candidateShape.Height = refShape.Height;

            if (pickupText &&
                refShape.HasTextFrame == MsoTriState.msoTrue &&
                candidateShape.HasTextFrame == MsoTriState.msoTrue)
            {
                var refParagraphCount = refShape.TextFrame2.TextRange.Paragraphs.Count;
                var candidateParagraphCount = candidateShape.TextFrame2.TextRange.Paragraphs.Count;
                var refTextRange = refShape.TextFrame.TextRange;
                var candidateTextRange = candidateShape.TextFrame.TextRange;

                for (var i = 1; i <= refParagraphCount && i <= candidateParagraphCount; i++)
                {
                    var refParagraph = refTextRange.Paragraphs(i);
                    var candidateParagraph = candidateTextRange.Paragraphs(i);
                    var candidateColor = candidateParagraph.Font.Color.RGB;

                    refParagraph.Copy();

                    var newCandidateRange = candidateParagraph.PasteSpecial();

                    newCandidateRange.Font.Color.RGB = candidateColor;
                }
            }

            if (pickupFormat)
            {
                refShape.PickUp();
                candidateShape.Apply();
            }
        }

        private static void SyncSingleAgendaBullet(PowerPointSlide refSlide, PowerPointSlide candidate)
        {
            if (refSlide == null || candidate == null)
            {
                return;
            }

            var refContentShape = refSlide.GetShapeWithName(PptLabsAgendaContentShapeName)[0];
            var candidateContentShape = candidate.GetShapeWithName(PptLabsAgendaContentShapeName)[0];

            SyncShape(refContentShape, candidateContentShape, false);
        }

        private static void SyncSingleAgendaVisual(PowerPointSlide refSlide, PowerPointSlide candidate)
        {
            if (refSlide == null || candidate == null)
            {
                return;
            }
        }

        private static void SyncSingleAgendaGeneral(PowerPointSlide refSlide, PowerPointSlide candidate)
        {
            // in this step, we should sync:
            // 1. Layout
            // 2. Design;
            // -3. Transition; -> this is no longer synced because each agenda may have different transition
            // 4. Shapes and their position, text;

            if (refSlide == null || candidate == null)
            {
                return;
            }

            candidate.Layout = refSlide.Layout;
            candidate.Design = refSlide.Design;
            candidate.Transition = refSlide.Transition;

            // syncronize extra shapes in reference slide
            var extraShapes = refSlide.Shapes.Cast<Shape>()
                                             .Where(shape => !candidate.HasShapeWithSameName(shape.Name))
                                             .Select(shape => shape.Name)
                                             .ToArray();

            if (extraShapes.Length != 0)
            {
                refSlide.Shapes.Range(extraShapes).Copy();
                candidate.Shapes.Paste();
            }

            // syncronize shapes position and size, except bullet content
            var sameShapes = refSlide.Shapes.Cast<Shape>()
                                            .Where(shape => shape.Name != PptLabsAgendaContentShapeName &&
                                                            candidate.HasShapeWithSameName(shape.Name));

            foreach (var refShape in sameShapes)
            {
                var candidateShape = candidate.GetShapeWithName(refShape.Name)[0];

                SyncShape(refShape, candidateShape);
            }
        }

        private static void UpdateColorScheme(PowerPointSlide slide, string section, bool isEnd)
        {
            var contentPlaceHolder = slide.GetShapeWithName(PptLabsAgendaContentShapeName)[0];
            var textRange = contentPlaceHolder.TextFrame.TextRange;
            var focusIndex = FindSectionIndex(section) - 1;
            var focusColor = isEnd ? _bulletDimColor : _bulletHighlightColor;

            RecolorTextRange(textRange, focusIndex, focusColor);
        }
        # endregion

        # region Event Handler
        private static void UpdateColorScheme(Color highlightColor, Color dimColor, Color defaultColor)
        {
            _bulletHighlightColor = highlightColor;
            _bulletDimColor = dimColor;
            _bulletDefaultColor = defaultColor;

            var sections = PowerPointPresentation.Current.Sections;
            var type = CurrentAgendaType;

            if (type != AgendaType.Bullet) return;

            // skip the default section
            foreach (var section in sections.Skip(1))
            {
                var startAgenda = FindSectionStartSlide(section, type);
                var endAgenda = FindSectionEndSlide(section, type);

                UpdateColorScheme(startAgenda, section, false);
                UpdateColorScheme(endAgenda, section, true);
            }
        }
        # endregion
    }
}
