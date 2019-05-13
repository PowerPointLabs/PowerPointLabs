using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace PowerPointLabs.ELearningLab.ELearningWorkspace.Model
{
    public class CustomItem: ClickItem, IEquatable<CustomItem>
    {
        public ObservableCollection<CustomSubItem> CustomItems { get; set; }

        public CustomItem(ObservableCollection<CustomSubItem> customSubItems)
        {
            CustomItems = customSubItems;
        }
        public override bool Equals(object other)
        {
            if (other == null || other.GetType() != GetType())
            {
                return false;
            }

            if (ReferenceEquals(other, this))
            {
                return true;
            }
            return Equals(other as CustomItem);
        }
        public bool Equals(CustomItem other)
        {
            return ClickNo == other.ClickNo && CustomItems.SequenceEqual(other.CustomItems);
        }

        public override int GetHashCode()
        {
            return -1125095958 + EqualityComparer<ObservableCollection<CustomSubItem>>.Default.GetHashCode(CustomItems);
        }
    }
}
