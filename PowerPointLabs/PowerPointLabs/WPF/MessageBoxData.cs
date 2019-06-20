using System.ComponentModel;
using PowerPointLabs.Utils;

namespace PowerPointLabs.WPF
{
    public class MessageBoxData : INotifyPropertyChanged
    {
        public MessageBoxData()
        {
            SetIcon(MessageBoxIcon.Asterisk);
        }

        public string IconSource { get; private set; }

        public event PropertyChangedEventHandler PropertyChanged;

        public void SetIcon(MessageBoxIcon icon)
        {
            IconSource = GetIcon(icon);
            OnPropertyChanged(nameof(IconSource));
        }
        private string GetIcon(MessageBoxIcon icon)
        {
            switch (icon)
            {
                case MessageBoxIcon.Asterisk:
                    return "..\\Resources\\About.png";
                case MessageBoxIcon.Error:
                    return "..\\Resources\\Help.png";
                case MessageBoxIcon.Exclamation:
                    return "..\\Resources\\Help.png";
                case MessageBoxIcon.Question:
                    return "..\\Resources\\Help.png";
                case MessageBoxIcon.None:
                default:
                    return "";
            }
        }
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

    }
}
