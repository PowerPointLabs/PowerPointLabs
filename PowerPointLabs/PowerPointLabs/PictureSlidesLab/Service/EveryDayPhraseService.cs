using System;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    class EveryDayPhraseService
    {
        private readonly string[] _phrases =
        {
            "A wise man speaks because he has something to say, a fool speaks because he has to say something. -- Plato",
            "It is better to say nothing and be thought a fool than to open your mouth and remove all doubt. -- Mark Twain",
            "It takes one hour of preparation for each minute of presentation time. -- Wayne Burgraff",
            "There are always three speeches, for every one you actually gave. The one you practiced, the one you gave, and the one you wish you gave. -- Dale Carnegie",
            "No one can remember more than three points. -- Philip Crosby",
            "The audience only pays attention as long as you know where you are going. -- Philip Crosby",
            "No one ever complains about a speech being too short! -- Ira Hayes",
            "Public speaking is not a talent – it is a skill. -- Unknown",
            "Top presenters have total control of their fears. They make fear their slave, not the master. -- Doug Malouf"
        };

        public string GetEveryDayPhrase()
        {
            return _phrases[new Random().Next(0, _phrases.Length)];
        }
    }
}
