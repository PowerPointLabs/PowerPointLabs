using System.ComponentModel;
using System.IO;
using System.Xml.Serialization;
using PowerPointLabs.Annotations;

namespace PowerPointLabs.ImageSearch.Model
{
    public class Notifiable : INotifyPropertyChanged
    {
        # region impl INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
        # endregion
    }
}
