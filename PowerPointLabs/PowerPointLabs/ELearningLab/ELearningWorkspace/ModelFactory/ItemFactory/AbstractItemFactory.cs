using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;
using PowerPointLabs.Models;

namespace PowerPointLabs.ELearningLab.ELearningWorkspace.ModelFactory
{
    public abstract class AbstractItemFactory
    {
        public IEnumerable<AbstractEffect> effects;

        public AbstractItemFactory(IEnumerable<AbstractEffect> effects)
        {
            this.effects = effects;
        }
        public ClickItem GetBlock()
        {
            return CreateBlock();
        }
        protected abstract ClickItem CreateBlock();
    }
}
