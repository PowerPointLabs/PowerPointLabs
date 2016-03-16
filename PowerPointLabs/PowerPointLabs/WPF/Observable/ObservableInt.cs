namespace PowerPointLabs.WPF.Observable
{
    public class ObservableInt : Model
    {
        private int _number;

        public int Number
        {
            get { return _number; }
            set
            {
                _number = value;
                OnPropertyChanged("Number");
            }
        }
    }
}
