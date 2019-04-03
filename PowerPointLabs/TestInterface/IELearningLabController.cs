using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestInterface
{
    public interface IELearningLabController
    {
        void OpenPane();
        void AddSelfExplanationItem();
        void Sync();
        void Reorder();
        void Delete();
    }
}
