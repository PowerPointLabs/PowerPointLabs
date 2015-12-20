namespace PowerPointLabs.WPF.Observable
{
    public class ObservableString : Model
    {
        private string _text;

        public string Text
        {
            get { return _text; }
            set { _text = value; OnPropertyChanged("Text"); }
        }
    }
}
