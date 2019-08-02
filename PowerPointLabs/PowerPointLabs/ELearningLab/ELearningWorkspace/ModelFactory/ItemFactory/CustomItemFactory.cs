using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;

namespace PowerPointLabs.ELearningLab.ELearningWorkspace.ModelFactory
{
    public class CustomItemFactory : AbstractItemFactory
    {
        public CustomItemFactory(IEnumerable<CustomEffect> effects):base(effects)
        { }
        protected override ClickItem CreateBlock()
        {
            if (effects.Count() == 0)
            {
                return null;
            }
            ObservableCollection<CustomSubItem> customItems = new ObservableCollection<CustomSubItem>();
            foreach (CustomEffect effect in effects)
            {
                customItems.Add(new CustomSubItem(effect.shapeName, effect.shapeId, effect.type));
            }
            return new CustomItem(customItems);
        }
    }
}
