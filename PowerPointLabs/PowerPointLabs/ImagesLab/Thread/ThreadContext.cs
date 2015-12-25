using System;
using System.Windows.Threading;
using PowerPointLabs.ImagesLab.Thread.Interface;

namespace PowerPointLabs.ImagesLab.Thread
{
    class ThreadContext : IThreadContext
    {
        private readonly Dispatcher _dispatcher;

        public ThreadContext(Dispatcher dispatcher)
        {
            _dispatcher = dispatcher;
        }

        public void Invoke(Action action)
        {
            _dispatcher.Invoke(action);
        }

        public void BeginInvoke(Action action)
        {
            _dispatcher.BeginInvoke(action);
        }
    }
}
