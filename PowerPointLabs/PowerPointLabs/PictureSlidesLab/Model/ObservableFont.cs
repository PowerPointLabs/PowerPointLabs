using System.Windows.Media;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    public class ObservableFont : WPF.Observable.Model
    {
        private FontFamily _font;

        public FontFamily Font
        {
            get { return _font; }
            set
            {
                _font = value;
                OnPropertyChanged("Font");
            }
        }
    }
}
