using System.ComponentModel;

namespace PowerPointLabs.PictureSlidesLab.Service.StylesWorker.Interface
{
    public interface IWorkerOrderMetadata
    {
        /// <summary>
        /// Define the execution order of workers.
        /// Also, it decides the z-index of the output from workers.
        /// Those workers executed at the beginning will have the output
        /// put at the back.
        /// </summary>
        [DefaultValue(int.MaxValue)]
        int WorkerOrder { get; }
    }
}
