namespace PowerPointLabs.AgendaLab.Templates
{
    internal class BulletAgendaTemplate : AgendaTemplate
    {
        public override Type Type
        {
            get { return Type.Bullet; }
        }

        public override void ConfigHead()
        {
            AgendaSlideConfig[] frontSlides = { };

            AgendaSlideConfig[] backSlides =
            {
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncEndBulletAgendaSlide, SlidePurpose.End),
            };

            AddConfiguration(frontSlides, backSlides);
        }

        public override void ConfigMiddle()
        {
            AgendaSlideConfig[] frontSlides =
            {
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncStartBulletAgendaSlide, SlidePurpose.Start),
            };

            AgendaSlideConfig[] backSlides =
            {
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncEndBulletAgendaSlide, SlidePurpose.End),
            };

            AddConfiguration(frontSlides, backSlides);
        }

        public override void ConfigEnd()
        {
            AgendaSlideConfig[] frontSlides =
            {
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncStartBulletAgendaSlide, SlidePurpose.Start),
            };

            AgendaSlideConfig[] backSlides =
            {
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncEndBulletAgendaSlide, SlidePurpose.End),
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncFinalBulletAgendaSlide, SlidePurpose.EndOfBulletAgenda),
            };

            AddConfiguration(frontSlides, backSlides);
        }
    }
}
