using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLabs.AgendaLab2
{
    internal enum Type
    {
        None,
        Bullet,
        Beam,
        Visual,
        Mixed
    };

    #region Slides
    internal enum SlidePurpose
    {
        None,
        Reference,

        // For Bullet Agenda
        Start,
        End,

        // For Visual Agenda
        ZoomIn,
        ZoomOut,
        FinalZoomOut,

        VisualAgendaSection,
        EndOfVisualAgenda
    };
    #endregion

    #region Shapes

    internal enum ShapePurpose
    {
        // For Bullet Agenda
        TitleShape,
        ContentShape,

        // For Beam Agenda
        BeamShapeMainGroup,
        BeamShapeText,
        BeamShapeBackground,
        BeamShapeHighlightedText,

        // For Visual Agenda
        VisualAgendaImage,
       
    }

    #endregion

    internal enum Direction
    {
        Top,
        Left,
        Bottom,
        Right
    };
}
