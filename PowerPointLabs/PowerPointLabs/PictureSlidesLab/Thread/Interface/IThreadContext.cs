using System;

namespace PowerPointLabs.PictureSlidesLab.Thread.Interface
{
    public interface IThreadContext
    {
        void Invoke(Action action);
        void BeginInvoke(Action action);
    }
}
