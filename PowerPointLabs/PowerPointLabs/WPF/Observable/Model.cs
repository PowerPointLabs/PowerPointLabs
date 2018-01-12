using System.ComponentModel;

namespace PowerPointLabs.WPF.Observable
{
    public abstract class Model : INotifyPropertyChanged
    {
        # region impl INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        # endregion
    }
}
