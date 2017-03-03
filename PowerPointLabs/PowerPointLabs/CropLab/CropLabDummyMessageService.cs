using System;

using PowerPointLabs.CustomControls;

namespace PowerPointLabs.CropLab
{
    internal class CropLabDummyMessageService : IMessageService
    {
        public void ShowErrorMessageBox(string content, Exception exception = null)
        {
            // do nothing in dummy
            return;
        }
    }
}
