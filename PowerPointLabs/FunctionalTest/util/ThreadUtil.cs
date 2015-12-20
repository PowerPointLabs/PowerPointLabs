using System.Threading;

namespace FunctionalTest.util
{
    class ThreadUtil
    {
        public static void WaitFor(int time)
        {
            var waitedFor = 0;
            while (waitedFor < time)
            {
                Thread.Sleep(time);
                waitedFor += time;
            }
        }
    }
}
