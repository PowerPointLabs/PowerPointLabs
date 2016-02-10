using System.ComponentModel;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface
{
    public interface IStyleOrderMetadata
    {
        /// <summary>
        /// Define the style order.
        /// </summary>
        [DefaultValue(int.MaxValue)]
        int StyleOrder { get; }
    }
}
