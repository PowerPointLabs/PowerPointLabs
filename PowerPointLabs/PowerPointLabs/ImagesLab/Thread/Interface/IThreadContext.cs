using System;

namespace PowerPointLabs.ImagesLab.Thread.Interface
{
    public interface IThreadContext
    {
        void Invoke(Action action);
        void BeginInvoke(Action action);
    }
}
