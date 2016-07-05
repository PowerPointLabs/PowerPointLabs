using System;
using System.ComponentModel;

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
