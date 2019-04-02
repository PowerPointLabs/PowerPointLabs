using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.ELearningLab.ELearningWorkspace.Model
{
    public abstract class AbstractEffect
    {
        public string shapeName;
        protected AbstractEffect(string shapeName)
        {
            this.shapeName = shapeName;
        }
    }
}
