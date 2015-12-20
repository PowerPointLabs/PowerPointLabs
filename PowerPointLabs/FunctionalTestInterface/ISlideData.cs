using System.Collections.Generic;

namespace FunctionalTestInterface
{
    public interface ISlideData
    {
        string SlideImage { get; }
        List<IEffectData> Animation { get; }
    }
}
