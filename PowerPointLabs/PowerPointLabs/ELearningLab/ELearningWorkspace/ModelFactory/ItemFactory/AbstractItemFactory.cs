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
        public IEnumerable<Effect> effects;
        public PowerPointSlide slide;

        public AbstractItemFactory(IEnumerable<Effect> effects, PowerPointSlide slide)
        {
            this.effects = effects;
            this.slide = slide;
        }
        public ClickItem GetBlock()
        {
            return CreateBlock();
        }
        protected abstract ClickItem CreateBlock();
    }
}
