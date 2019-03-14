using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.ELearningLab.ELearningWorkspace.Model
{
    public class CustomClickItem: ClickItem, IEquatable<CustomClickItem>
    {
        public ObservableCollection<CustomSubItem> CustomItems { get; set; }

        public CustomClickItem(ObservableCollection<CustomSubItem> customSubItems)
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
            return Equals(other as CustomClickItem);
        }
        public bool Equals(CustomClickItem other)
        {
            return ClickNo == other.ClickNo && CustomItems.SequenceEqual(other.CustomItems);
        }

        public override int GetHashCode()
        {
            return -1125095958 + EqualityComparer<ObservableCollection<CustomSubItem>>.Default.GetHashCode(CustomItems);
        }
    }
}
