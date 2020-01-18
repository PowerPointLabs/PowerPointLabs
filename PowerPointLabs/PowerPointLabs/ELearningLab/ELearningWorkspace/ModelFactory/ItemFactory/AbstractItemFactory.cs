using System.Collections.Generic;

using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;

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
