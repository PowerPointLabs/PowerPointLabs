using System;
using System.Runtime.InteropServices;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.AgendaLab
{
    /// <summary>
    /// AgendaShape class for Encoding/Decoding shape names
    /// An AgendaShape object carries shape metadata.
    /// </summary>
    internal class AgendaShape
    {
        public readonly ShapePurpose ShapePurpose;
        public readonly AgendaSection Section;

        private static readonly string[] Delim = { "_&^@" };
        private const string Prefix = "PptLabsAgenda";

        private AgendaShape(ShapePurpose shapePurpose, AgendaSection section)
        {
            ShapePurpose = shapePurpose;
            Section = section;
        }

        /// <summary>
        /// Universal Encode function used for all Agenda Shapes.
        /// Packs a set of agenda shape properties to a shape name. Paired with the Decode function.
        /// </summary>
        public static string Encode(ShapePurpose shapePurpose, AgendaSection section)
        {
            string[] parameters = { Prefix, shapePurpose.ToString(), section.Encode() };
            return string.Join(Delim[0], parameters);
        }

        /// <summary>
        /// Universal Encode function used for all Agenda Shapes.
        /// Unpacks a shape name into a set of shape properties. Paired with the Encode function.
        /// </summary>
        public static AgendaShape Decode(string shapeName)
        {
            string[] parameters = shapeName.Split(Delim, StringSplitOptions.None);

            if (parameters.Length != 3)
            {
                return null;
            }

            if (parameters[0] != Prefix)
            {
                return null;
            }

            ShapePurpose shapePurpose;
            if (!Enum.TryParse(parameters[1], out shapePurpose))
            {
                return null;
            }

            AgendaSection section = AgendaSection.Decode(parameters[2]);

            return new AgendaShape(shapePurpose, section);
        }

        public static bool IsBeamShape(Shape shape)
        {
            AgendaShape agendaShape = Decode(shape);
            if (agendaShape == null)
            {
                return false;
            }

            return agendaShape.ShapePurpose == ShapePurpose.BeamShapeMainGroup;
        }

        public static AgendaShape Decode(Shape shape)
        {
            if (shape == null)
            {
                return null;
            }

            try
            {
                return Decode(shape.Name);
            }
            catch (COMException)
            {
                // sometims the shape is inaccessible (perhaps deleted. never occurred to me before.)
                // in this case, a COMException is thrown. so we return null.
                return null;
            }
        }


        /// <summary>
        /// </summary>
        /// <param name="condition">Input a condition on (AgendaShape : bool)</param>
        /// <returns>Output a condition on (Shape : bool).</returns>
        public static Func<Shape, bool> MeetsConditions(Predicate<AgendaShape> condition)
        {
            return shape =>
            {
                AgendaShape agendaShape = Decode(shape);
                if (agendaShape == null)
                {
                    return false;
                }

                return condition(agendaShape);
            };
        }


        /// <summary>
        /// Returns a condition (function) that is true iff the shape's purpose is the purpose specified in the argument.
        /// </summary>
        public static Func<Shape, bool> WithPurpose(ShapePurpose purpose)
        {
            return shape =>
            {
                AgendaShape agendaShape = Decode(shape);
                if (agendaShape == null)
                {
                    return false;
                }

                return agendaShape.ShapePurpose == purpose;
            };
        }


        /// <summary>
        /// Stores metadata in the shape by setting its name.
        /// </summary>
        public static void SetShapeName(Shape shape, ShapePurpose shapePurpose, AgendaSection section)
        {
            shape.Name = Encode(shapePurpose, section);
        }


        public static bool IsAnyAgendaShape(Shape shape)
        {
            return Decode(shape) != null;
        }
    }
}
