using System;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using PowerPointLabs.Models;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    internal static class AgendaLab
    {
        private const string PptLabsAgendaSlideNameFormat = "PptLabs{0}Agenda{1}Slide {2}";
        private const string PptLabsAgendaTitleShapeName = "PptLabsAgendaTitle";
        private const string PptLabsAgendaContentShapeName = "PptLabsAgendaContent";
        private const string PptLabsAgendaSlideTypeSearchPattern = @"PptLabs(\w+)Agenda(?:Start|End)Slide";

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
                var agendaSlideSearchPattern = new Regex(PptLabsAgendaSlideTypeSearchPattern);

                return PowerPointPresentation.Current.Slides.Any(slide => agendaSlideSearchPattern.IsMatch(slide.Name));
            }
        }
        # endregion

        # region API
        public static void GenerateAgenda(AgendaType type)
        {
            switch (type)
            {
                case AgendaType.Beam:
                    GenerateBeamAgenda();
                    break;
                case AgendaType.Bullet:
                    GenerateBulletAgenda();
                    break;
                case AgendaType.Visual:
                    GenerateVisualAgenda();
                    break;
            }
        }

        public static void RemoveAgenda()
        {
            var slides = PowerPointPresentation.Current.Slides;
            var generatedAgendaNamePattern = new Regex(PptLabsAgendaSlideTypeSearchPattern);

            foreach (var slide in slides)
            {
                if (generatedAgendaNamePattern.IsMatch(slide.Name))
                {
                    slide.Delete();
                }
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
        private static void AddAgendaSlide(AgendaType type, string section, bool isEnd)
        {
            var sectionIndex = FindSectionIndex(section);
            var sectionEndIndex = FindSectionEnd(section);

            var newSlide =
                PowerPointSlide.FromSlideFactory(PowerPointPresentation.Current
                                                                       .Presentation
                                                                       .Slides
                                                                       .Add(isEnd ? sectionEndIndex + 1 : 1,
                                                                            PpSlideLayout.ppLayoutText));

            newSlide.Name = string.Format(PptLabsAgendaSlideNameFormat, type, isEnd ? "End" : "Start", section);
            newSlide.Transition.EntryEffect = PpEntryEffect.ppEffectFadeSmoothly;
            newSlide.Transition.Duration = 0.25f;

            newSlide.Shapes.Placeholders[1].Name = PptLabsAgendaTitleShapeName;
            newSlide.Shapes.Placeholders[2].Name = PptLabsAgendaContentShapeName;

            if (!isEnd)
            {
                newSlide.GetNativeSlide().MoveToSectionStart(sectionIndex);
            }

            switch (type)
            {
                case AgendaType.Bullet:
                    AddAgendaSlideBulletType(newSlide, section, isEnd);
                    break;
                case AgendaType.Visual:
                    break;
            }
        }

        private static void AddAgendaSlideBulletType(PowerPointSlide slide, string section, bool isEnd)
        {
            // set title
            slide.Shapes.Placeholders[1].TextFrame.TextRange.Text = "Agenda";

            // set agenda content
            var contentPlaceHolder = slide.Shapes.Placeholders[2];
            var textRange = contentPlaceHolder.TextFrame.TextRange;

            // since we skip the default section, relative section index should be substracted by 1
            var relativeSectionIndex = FindSectionIndex(section) - 1;

            textRange.Text = _agendaText;

            for (var i = 1; i < relativeSectionIndex; i++)
            {
                textRange.Paragraphs(i).Font.Color.RGB = PowerPointLabsGlobals.CreateRGB(Color.Gray);
            }

            textRange.Paragraphs(relativeSectionIndex).Font.Color.RGB =
                PowerPointLabsGlobals.CreateRGB(isEnd ? Color.Gray : Color.Red);
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

            var typeSearchRegex = new Regex(PptLabsAgendaSlideTypeSearchPattern);

            return (AgendaType)Enum.Parse(typeof(AgendaType), typeSearchRegex.Match(startRef.Name).Groups[1].Value);
        }

        private static void GenerateBeamAgenda()
        {
            throw new NotImplementedException();
        }

        private static void GenerateBulletAgenda()
        {
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

            // need to use '\r' as paragraph indicator, not '\n'!
            _agendaText = sections.Aggregate((current, next) => current + "\r" + next) + "\r";

            foreach (var section in sections)
            {
                AddAgendaSlide(AgendaType.Bullet, section, false);
                AddAgendaSlide(AgendaType.Bullet, section, true);
            }
        }

        private static void GenerateVisualAgenda()
        {
            throw new NotImplementedException();
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
                                                             shape.Name != PptLabsAgendaContentShapeName)
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
