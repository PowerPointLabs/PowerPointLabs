using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public class GenericEventArgs<T> : EventArgs
    {
        public GenericEventArgs(T eventData, string filepath)
        {
            this.EventData = eventData;
            this.FilePath = filepath;
        }

        public GenericEventArgs(T eventData)
        {
            this.EventData = eventData;
            this.FilePath = null;
        }

        public T EventData { get; private set; }

        public string FilePath { get; private set; }
    }
}
