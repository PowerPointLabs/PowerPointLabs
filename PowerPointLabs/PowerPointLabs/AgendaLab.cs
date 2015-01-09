using System;
using System.Collections.Generic;
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

        public static void SyncrhonizeAgenda()
        {
            // find the agenda for the first section as reference
            PowerPointSlide startRef;
            PowerPointSlide endRef;

            var type = FindSyncReference(out startRef, out endRef);

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
            var sectionIndex = FindSectionIndex(section) + 1;
            var sectionEndIndex = FindSectionEnd(section);

            var newSlide =
                PowerPointSlide.FromSlideFactory(
                    PowerPointCurrentPresentationInfo.CurrentPresentation.Slides.Add(isEnd ? sectionEndIndex : 1,
                                                                                     PpSlideLayout.ppLayoutText));

            newSlide.Name = string.Format(PptLabsAgendaSlideNameFormat, type, isEnd ? "Start" : "End", section);
            newSlide.Transition.EntryEffect = PpEntryEffect.ppEffectFadeSmoothly;
            newSlide.Transition.Duration = 0.25f;
            
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

            // since we skip the default section, by right relative section index should be substracted
            // by 1, but FindSectionIndex returns a 0-based number, it's ok to use the number without
            // doing substraction
            var relativeSectionIndex = FindSectionIndex(section);

            textRange.Text = _agendaText;

            for (var i = 1; i < relativeSectionIndex; i ++ )
            {
                textRange.Paragraphs(i).Font.Color.RGB = PowerPointLabsGlobals.CreateRGB(Color.Gray);
            }

            textRange.Paragraphs(relativeSectionIndex).Font.Color.RGB =
                PowerPointLabsGlobals.CreateRGB(isEnd ? Color.Gray : Color.Red);
        }

        private static void SyncAgendaBulletType(PowerPointSlide startRef, PowerPointSlide endRef)
        {
            throw new NotImplementedException();
        }

        private static int FindSectionEnd(string section)
        {
            var sectionIndex = FindSectionIndex(section) + 1;

            return FindSectionEnd(sectionIndex);
        }

        private static int FindSectionEnd(int sectionIndex)
        {
            var sectionProperties = PowerPointCurrentPresentationInfo.CurrentPresentation.SectionProperties;

            return sectionProperties.FirstSlide(sectionIndex) + sectionProperties.SlidesCount(sectionIndex);
        }

        private static int FindSectionStart(string section)
        {
            var sectionIndex = FindSectionIndex(section) + 1;

            return FindSectionStart(sectionIndex);
        }

        private static int FindSectionStart(int sectionIndex)
        {
            var sectionProperties = PowerPointCurrentPresentationInfo.CurrentPresentation.SectionProperties;

            return sectionProperties.FirstSlide(sectionIndex);
        }

        private static int FindSectionIndex(string section)
        {
            return PowerPointCurrentPresentationInfo.Sections.FindIndex(name => name == section);
        }

        private static AgendaType FindSyncReference(out PowerPointSlide startRef, out PowerPointSlide endRef)
        {
            var curPresentation = PowerPointCurrentPresentationInfo.CurrentPresentation;

            // the first meaningful section is the second section
            startRef = PowerPointSlide.FromSlideFactory(curPresentation.Slides[FindSectionStart(2)]);
            endRef = PowerPointSlide.FromSlideFactory(curPresentation.Slides[FindSectionEnd(2)]);

            var typeSearchRegex = new Regex(PptLabsAgendaSlideTypeSearchPattern);

            return (AgendaType) Enum.Parse(typeof(AgendaType), typeSearchRegex.Match(startRef.Name).Groups[1].Value);
        }

        private static void GenerateBeamAgenda()
        {
            throw new NotImplementedException();
        }

        private static void GenerateBulletAgenda()
        {
            var sections = PowerPointCurrentPresentationInfo.Sections.Skip(1).ToList();

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
            _agendaText = sections.Aggregate((current, next) => current + "\r" + next);

            Globals.ThisAddIn.Application.StartNewUndoEntry();

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
        # endregion
    }
}
