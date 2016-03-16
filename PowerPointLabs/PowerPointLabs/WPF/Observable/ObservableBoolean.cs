namespace PowerPointLabs.WPF.Observable
{
    public class ObservableBoolean : Model
    {
        private bool _flag;

        public bool Flag
        {
            get { return _flag; }
            set
            {
                _flag = value;
                OnPropertyChanged("Flag");
            }
        }
    }
}
