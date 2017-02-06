namespace PowerPointLabs.AgendaLab.Templates
{
    internal class VisualAgendaTemplate : AgendaTemplate
    {
        public override Type Type
        {
            get { return Type.Visual; }
        }

        public override void ConfigHead()
        {
            AgendaSlideConfig[] frontSlides = { };
            AgendaSlideConfig[] backSlides = { };
            AddConfiguration(frontSlides, backSlides);
        }

        public override void ConfigMiddle()
        {
            AgendaSlideConfig[] frontSlides =
            {
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncVisualAgendaSlide, SlidePurpose.VisualAgendaSection),
            };

            AgendaSlideConfig[] backSlides = { };

            AddConfiguration(frontSlides, backSlides);
        }

        public override void ConfigEnd()
        {
            AgendaSlideConfig[] frontSlides =
            {
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncVisualAgendaSlide, SlidePurpose.VisualAgendaSection),
            };

            AgendaSlideConfig[] backSlides =
            {
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncVisualAgendaEndSlide, SlidePurpose.EndOfVisualAgenda),
            };

            AddConfiguration(frontSlides, backSlides);
        }
    }
}
