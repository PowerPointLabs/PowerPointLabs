using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.ShapesLab.Views
{
    class CustomComboBoxItem
    {
        public string actualName;
        public bool isDefaultCategory;

        private const string DefaultCategorySuffix = " (default)";

        public CustomComboBoxItem(string name, string defaultCategory)
        {
            actualName = name;
            isDefaultCategory = name == defaultCategory;
        }

        public void SetNewDefaultCategory(string defaultCategory)
        {
            isDefaultCategory = actualName == defaultCategory;
        }

        override public string ToString()
        {
            return actualName + (isDefaultCategory ? DefaultCategorySuffix : "");
        }
    }
}
