using System.ComponentModel;

namespace PowerPointLabs.ImageSearch.Model
{
    public class Notifiable : INotifyPropertyChanged
    {
        # region impl INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
        # endregion
    }
}
