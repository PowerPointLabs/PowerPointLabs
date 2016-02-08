using System.Collections.Generic;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker
{
    class StyleWorkerFactory
    {
        public static IList<IStyleWorker> GetAllStyleWorkers()
        {
            return new List<IStyleWorker>
            {
                new TextStyleWorker(),
                new StyleEmbeddingWorker(),
                new OverlayStyleWorker(),
                new BlurStyleWorker(),
                new BannerStyleWorker(),
                new TextBoxStyleWorker(),
                new OutlineStyleWorker(),
                new FrameStyleWorker(),
                new CircleStyleWorker(),
                new TriangleStyleWorker(),
                new PictureCitationStyleWorker()
            };
        }
    }
}
