using System;

namespace PowerPointLabs.ResizeLab
{
    public interface IResizeLabPane
    {
        void ShowErrorMessageBox(string content, Exception exception = null);
    }
}