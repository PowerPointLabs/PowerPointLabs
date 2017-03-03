using System;

namespace PowerPointLabs.CustomControls
{
    internal interface IMessageService
    {
        void ShowErrorMessageBox(string content, Exception exception = null);
    }

    internal static class MessageServiceFactory
    {
        private static IMessageService cropLabMessageService;

        public static IMessageService GetCropLabMessageService()
        {
            if (cropLabMessageService == null)
            {
                cropLabMessageService = new CropLab.CropLabMessageService();
            }
            return cropLabMessageService;
        }

        public static IMessageService GetCropLabDummyMessageService()
        {
            return new CropLab.CropLabDummyMessageService();
        }
    }

}
