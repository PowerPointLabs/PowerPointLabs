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

namespace PowerPointLabs.AgendaLab
{
    internal static partial class AgendaLabMain
    {
#pragma warning disable 0618
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
                {
                    return new List<PowerPointSlide>();
                }

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
                if (slides.Count == 0)
                {
                    return null;
                }
                return slides[0];
            }
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
            if (referenceSlide == null)
            {
                return Type.None;
            }
            return AgendaSlide.Decode(referenceSlide).AgendaType;
        }

        private static Type GetAnyAgendaSlideType()
        {
            var agendaSlide = PowerPointPresentation.Current
                                                    .Slides
                                                    .FirstOrDefault(slide => AgendaSlide.IsAnyAgendaSlide(slide) ||
                                                                             HasBeamShape(slide));
            if (agendaSlide == null)
            {
                return Type.None;
            }

            if (HasBeamShape(agendaSlide))
            {
                return Type.Beam;
            }

            return AgendaSlide.Decode(agendaSlide).AgendaType;
        }

        /// <summary>
        /// Only checks the first slide. Checks if it is a suitable candidate for the reference slide, and if it is,
        /// returns it. Else returns null.
        /// </summary>
        private static PowerPointSlide TryFindSuitableRefSlide(Type type)
        {
            var firstSlide = PowerPointPresentation.Current.FirstSlide;
            if (firstSlide == null)
            {
                return null;
            }

            if (InvalidReferenceSlide(type, firstSlide))
            {
                return null;
            }

            return firstSlide;
        }

        /// <summary>
        /// The list is 0-indexed, section indexes are 1-indexed.
        /// </summary>
        private static List<AgendaSection> Sections
        {
            get
            {
                var sectionProperty = PowerPointPresentation.Current.SectionProperties;
                var sections = new List<AgendaSection>(sectionProperty.Count);

                for (var i = 1; i <= sectionProperty.Count; i++)
                {
                    sections.Add(new AgendaSection(sectionProperty.Name(i), i));
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


        private static int NumberOfSections
        {
            get { return PowerPointPresentation.Current.Sections.Count; }
        }

        private static PowerPointSlide FindSectionFirstNonAgendaSlide(int sectionIndex)
        {
            var presentation = PowerPointPresentation.Current;
            int slideCount = presentation.SlideCount;

            int currentIndex = SectionFirstSlideIndex(sectionIndex);
            var slide = presentation.GetSlide(currentIndex);

            while (AgendaSlide.IsAnyAgendaSlide(slide))
            {
                currentIndex++;
                if (currentIndex > slideCount)
                {
                    return null;
                }

                slide = presentation.GetSlide(currentIndex);
            }
            return slide;
        }

        private static PowerPointSlide FindSectionLastNonAgendaSlide(int sectionIndex)
        {
            var presentation = PowerPointPresentation.Current;

            int currentIndex = SectionLastSlideIndex(sectionIndex);
            var slide = presentation.GetSlide(currentIndex);

            while (AgendaSlide.IsAnyAgendaSlide(slide))
            {
                currentIndex--;
                if (currentIndex <= 0)
                {
                    return null;
                }

                slide = presentation.GetSlide(currentIndex);
            }
            return slide;
        }

        /// <summary>
        /// Assumes that there are at least two sections.
        /// </summary>
        private static List<PowerPointSlide> AllSlidesAfterFirstSection()
        {
            var slides = PowerPointPresentation.Current.Slides;
            int firstSlideIndex = SectionFirstSlideIndex(2);
            int lastSlideIndex = PowerPointPresentation.Current.SlideCount;

            return slides.GetRange(firstSlideIndex - 1, lastSlideIndex - firstSlideIndex + 1);
        }

        /// <summary>
        /// 1-indexed.
        /// </summary>
        private static PowerPointSlide TryGetSlideAtIndex(int index)
        {
            if (index > PowerPointPresentation.Current.SlideCount)
            {
                return null;
            }

            if (index < 1)
            {
                return null;
            }

            return PowerPointPresentation.Current.GetSlide(index);
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

        private static AgendaSection GetSlideSection(PowerPointSlide slide)
        {
            var slideIndex = slide.Index;

            var sections = Sections;
            var firstSlideIndexes = sections.Select(SectionFirstSlideIndex).ToList();

            int i = 0;
            for (; i < sections.Count; ++i)
            {
                if (slideIndex < firstSlideIndexes[i])
                {
                    break;
                }
            }
            return sections[i - 1];
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

            if (lastSlideIndex <= -1)
            {
                lastSlideIndex = -1;
            }
            return lastSlideIndex;
        }

        /// <summary>
        /// 0-indexed.
        /// </summary>
        private static List<PowerPointSlide> GetSectionSlides(AgendaSection section)
        {
            var slides = PowerPointPresentation.Current.Slides;

            int firstSlideIndex = SectionFirstSlideIndex(section);
            int lastSlideIndex = SectionLastSlideIndex(section);
            if (firstSlideIndex == -1 || lastSlideIndex == -1)
            {
                return new List<PowerPointSlide>();
            }

            return slides.GetRange(firstSlideIndex - 1, lastSlideIndex - firstSlideIndex + 1);
        }

        private static bool IsLastSection(int sectionIndex)
        {
            return sectionIndex == NumberOfSections;
        }

        #endregion
    }
}
