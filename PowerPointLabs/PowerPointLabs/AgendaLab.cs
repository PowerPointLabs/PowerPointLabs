using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PowerPointLabs.Models;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    internal static class AgendaLab
    {
        private const string PptLabsAgendaSlideNameFormat = "PptLabs{0}Agenda{1}Slide {2}";

        private static string _agendaText;

        # region Enum
        public enum AgendaType
        {
            Bullet,
            Visual
        };
        # endregion

        # region API
        public static void GenerateAgenda(AgendaType type)
        {
            switch (type)
            {
                case AgendaType.Bullet:
                    GenerateBulletAgenda();
                    break;
                case AgendaType.Visual:
                    GenerateVisualAgenda();
                    break;
            }
        }

        public static void UpdateAgenda()
        {
            throw new NotImplementedException();
        }
        # endregion

        # region Helper Functions
        private static void AddAgendaSlide(AgendaType type, string section, bool isEnd)
        {
            var sectionIndex = FindSectionIndex(section) + 1;
            var sectionEndIndex = FindSectionEnd(section);

            var newSlide = PowerPointSlide.FromSlideFactory(Globals.ThisAddIn
                                                                   .Application
                                                                   .ActivePresentation
                                                                   .Slides.Add(isEnd ? sectionEndIndex : 1,
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

        private static int FindSectionEnd(string section)
        {
            var sectionProperties = PowerPointCurrentPresentationInfo.CurrentPresentation.SectionProperties;
            var sectionIndex = FindSectionIndex(section) + 1;

            return sectionProperties.FirstSlide(sectionIndex) + sectionProperties.SlidesCount(sectionIndex);
        }

        private static int FindSectionIndex(string section)
        {
            return PowerPointCurrentPresentationInfo.Sections.FindIndex(name => name == section);
        }

        private static void GenerateBulletAgenda()
        {
            // need to use '\r' as paragraph indicator, not '\n'!
            var sections = PowerPointCurrentPresentationInfo.Sections.Skip(1).ToList();
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
