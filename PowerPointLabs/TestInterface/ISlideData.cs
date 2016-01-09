using System.Collections.Generic;

namespace TestInterface
{
    public interface ISlideData
    {
        string SlideImage { get; }
        List<IEffectData> Animation { get; }
    }
}
