
namespace PowerPointLabs.AgendaLab
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
        EndOfBulletAgenda,

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
