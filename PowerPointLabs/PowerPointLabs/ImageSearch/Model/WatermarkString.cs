namespace PowerPointLabs.ImageSearch.Model
{
    public class WatermarkString : Notifiable
    {
        private string _watermark;

        public string Watermark
        {
            get
            {
                return _watermark;
            }
            set
            {
                _watermark = value;
                OnPropertyChanged("Watermark");
            }
        }
    }
}
