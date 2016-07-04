using System;
using System.ComponentModel;
using System.IO;
using System.Xml.Serialization;
using PowerPointLabs.ActionFramework.Common.Log;
using System.Linq;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    /// <summary>
    /// StyleOption provides settings or options for a style.
    /// 
    /// To support any new properties, create a new partial class of 
    /// StyleOption under folder `StyleOption.Partial`.
    /// </summary>
    [Serializable]
    public partial class StyleOption
    {
        public StyleOption()
        {
            Init();
        }

        public override bool Equals(object obj)
        {
            var thisType = GetType();
            var objType = obj.GetType();

            foreach (var propertyInfo in thisType.GetProperties().Where(p => !p.Name.Contains("StyleName")))
            {
                if (propertyInfo != objType.GetProperty(propertyInfo.Name))
                {
                    return false;
                }
            }

            return true;
        }

        public override int GetHashCode()
        {
            return GetHashCode();
        }

        #region Initialization
        private void Init()
        {
            foreach (PropertyDescriptor property in TypeDescriptor.GetProperties(this))
            {
                var myAttribute = (DefaultValueAttribute) property
                    .Attributes[typeof(DefaultValueAttribute)];
                if (myAttribute != null)
                {
                    property.SetValue(this, myAttribute.Value);
                }
            }
        }
        #endregion
    }
}
