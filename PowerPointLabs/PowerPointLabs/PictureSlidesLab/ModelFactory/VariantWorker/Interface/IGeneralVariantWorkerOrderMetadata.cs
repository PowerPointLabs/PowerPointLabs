using System.ComponentModel;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.VariantWorker.Interface
{
    public interface IGeneralVariantWorkerOrderMetadata
    {
        /// <summary>
        /// Define the general variant worker order.
        /// </summary>
        [DefaultValue(int.MaxValue)]
        int GeneralVariantWorkerOrder { get; }
    }
}
