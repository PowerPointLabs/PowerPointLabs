using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;

namespace PowerPointLabs.AgendaLab
{
    /// <summary>
    /// AgendaSlide class for Encoding/Decoding slide names
    /// An AgendaSlide object carries slide metadata.
    /// </summary>
    internal class AgendaSlide
    {
        public readonly Type AgendaType;
        public readonly SlidePurpose SlidePurpose;
        public readonly AgendaSection Section;

        private static readonly string[] Delim = { "_&^@" };
        private const string Prefix = "PptLabsAgenda";

        private AgendaSlide(Type type, SlidePurpose slidePurpose, AgendaSection section)
        {
            AgendaType = type;
            SlidePurpose = slidePurpose;
            Section = section;
        }

        /// <summary>
        /// Universal Encode function used for all Agenda Slides.
        /// Packs a set of agenda slide properties to a slide name. Paired with the Decode function.
        /// </summary>
        public static string Encode(Type agendaType, SlidePurpose slidePurpose, AgendaSection section)
        {
            string[] parameters = { Prefix, agendaType.ToString(), slidePurpose.ToString(), section.Encode() };
            return string.Join(Delim[0], parameters);
        }

        /// <summary>
        /// Universal Encode function used for all Agenda Slides.
        /// Unpacks a slide name into a set of slide properties. Paired wit hthe Encode function.
        /// </summary>
        public static AgendaSlide Decode(string slide)
        {
            string[] parameters = slide.Split(Delim, StringSplitOptions.None);

            if (parameters.Length != 4)
            {
                return null;
            }

            if (parameters[0] != Prefix)
            {
                return null;
            }

            Type type;
            SlidePurpose slidePurpose;
            if (!Enum.TryParse(parameters[1], out type))
            {
                return null;
            }

            if (!Enum.TryParse(parameters[2], out slidePurpose))
            {
                return null;
            }

            AgendaSection section = AgendaSection.Decode(parameters[3]);

            return new AgendaSlide(type, slidePurpose, section);
        }

        public static AgendaSlide Decode(PowerPointSlide slide)
        {
            if (slide == null)
            {
                return null;
            }

            try
            {
                return Decode(slide.Name);
            }
            catch (COMException)
            {
                // sometims the shape is inaccessible (perhaps deleted. never occurred to me before.)
                // in this case, a COMException is thrown. so we return null.
                return null;
            }
        }

        public static AgendaSlide Decode(Slide slide)
        {
            if (slide == null)
            {
                return null;
            }

            try
            {
                return Decode(slide.Name);
            }
            catch (COMException)
            {
                // sometims the shape is inaccessible (perhaps deleted. never occurred to me before.)
                // in this case, a COMException is thrown. so we return null.
                return null;
            }
        }

        /// <summary>
        /// Stores metadata in the slide by setting its name.
        /// </summary>
        public static void SetSlideName(PowerPointSlide slide, Type agendaType, SlidePurpose slidePurpose, AgendaSection section)
        {
            slide.Name = Encode(agendaType, slidePurpose, section);
        }

        /// <summary>
        /// Assigns the slide's section to None, which is encoded as an arbitrary string of numbers that is guaranteed to be unique.
        /// Used to prevent duplicate slide names, which crashes powerpoint.
        /// </summary>
        public static void AssignUniqueSectionName(PowerPointSlide slide)
        {
            var properties = Decode(slide);
            slide.Name = Encode(properties.AgendaType, properties.SlidePurpose, AgendaSection.None);
        }

        /// <summary>
        /// Sets the name to that used by the reference (template) slide.
        /// </summary>
        public static void SetAsReferenceSlideName(PowerPointSlide slide, Type type)
        {
            slide.Name = Encode(type, SlidePurpose.Reference, AgendaSection.None);
        }

        public static bool IsReferenceslide(PowerPointSlide slide)
        {
            var agendaSlide = Decode(slide);
            if (agendaSlide == null)
            {
                return false;
            }

            return agendaSlide.SlidePurpose == SlidePurpose.Reference;
        }

        public static bool IsReferenceslide(Slide slide)
        {
            var agendaSlide = Decode(slide);
            if (agendaSlide == null)
            {
                return false;
            }

            return agendaSlide.SlidePurpose == SlidePurpose.Reference;
        }

        // convenience method.
        public static bool IsNotReferenceslide(PowerPointSlide slide)
        {
            return !IsReferenceslide(slide);
        }

        // convenience method.
        public static bool IsNotReferenceslide(Slide slide)
        {
            return !IsReferenceslide(slide);
        }

        /// <summary>
        /// Duplicate function of MeetsConditions2 to return a function (PowerPointSlide => bool) instead of (Slide => bool).
        /// 
        /// Example Usage:
        /// <code>
        ///     var condition = AgendaSlide.MeetsConditions(
        ///            agendaSlide =>
        ///                agendaSlide.AgendaType == Type.Visual &&
        ///                agendaSlide.SlidePurpose == SlidePurpose.VisualAgendaSection);
        ///     return powerPointSlides.Where(condition);
        /// </code>
        /// </summary>
        /// <param name="condition">Input a condition on (AgendaSlide : bool)</param>
        /// <returns>Output a condition on (PowerPointSlide : bool).</returns>
        public static Func<PowerPointSlide, bool> MeetsConditions(Predicate<AgendaSlide> condition)
        {
            return slide =>
            {
                var agendaSlide = Decode(slide);
                if (agendaSlide == null)
                {
                    return false;
                }

                return condition(agendaSlide);
            };
        }

        /// <summary>
        /// Same as MeetConditions, but returns a function from Slide to bool instead.
        /// </summary>
        /// <param name="condition">Input a condition on (AgendaSlide : bool)</param>
        /// <returns>Output a condition on (Slide : bool).</returns>
        public static Func<Slide, bool> MeetsConditions2(Predicate<AgendaSlide> condition)
        {
            return slide =>
            {
                var agendaSlide = Decode(slide);
                if (agendaSlide == null)
                {
                    return false;
                }

                return condition(agendaSlide);
            };
        }

        public static bool MatchesPurpose(PowerPointSlide slide, SlidePurpose purpose)
        {
            var agendaSlide = Decode(slide);
            if (agendaSlide == null)
            {
                return false;
            }

            return agendaSlide.SlidePurpose == purpose;
        }

        public static bool IsAnyAgendaSlide(PowerPointSlide slide)
        {
            return Decode(slide) != null;
        }

        public static bool IsAnyAgendaSlide(Slide slide)
        {
            return Decode(slide) != null;
        }
    }
}
