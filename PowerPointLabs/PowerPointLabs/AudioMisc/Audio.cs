using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.AudioMisc
{
    public class Audio
    {

        public enum AudioType
        {
            Record,
            Auto,
            Unrecognized
        }

        public const int GeneratedSamplingRate = 22050;
        public const int RecordedSamplingRate = 11025;
        public const int GeneratedBitRate = 16;
        public const int RecordedBitRate = 16;
        public const int GeneratedChannels = 1;
        public const int RecordedChannels = 1;
        public const string RecordedFormatName = "WAVE";
        public const string RecordedFormatExtension = "wav";

        public string Name { get; set; }
        public int MatchScriptID { get; set; }
        public string SaveName { get; set; }
        public string Length { get; set; }
        public int LengthMillis { get; set; }
        public AudioType Type { get; set; }

        // Tag key to store audio type in shape
        private const string AudioTypeTagName = "AUDIO_TAG";

        /// <summary>
        /// Default constructor.
        /// </summary>
        public Audio() {}


        /// <summary>
        /// Initialize Audio from sound file
        /// </summary>
        public Audio(string name, string saveName, int matchScriptID)
        {
            Name = name;
            MatchScriptID = matchScriptID;
            SaveName = saveName;
            LengthMillis = AudioHelper.GetAudioLength(saveName);
            Length = AudioHelper.ConvertMillisToTime(LengthMillis);
            Type = AudioHelper.GetAudioType(saveName);
        }

        /// <summary>
        /// Initialize Audio from a sound shape
        /// </summary>
        public Audio(Shape shape, string saveName)
        {
            // detect audio type from shape tag
            AudioType audioType = GetShapeAudioType(shape);
            this.Type = audioType;

            // derive matched id from shape name
            string[] temp = shape.Name.Split(new[] { ' ' });
            if (temp.Length < 3)
            {
                throw new FormatException(NarrationsLabText.RecorderUnrecognizeAudio);
            }
            this.MatchScriptID = Int32.Parse(temp[2]);

            // get corresponding audio
            this.Name = shape.Name;
            this.SaveName = saveName;
            this.Length = AudioHelper.GetAudioLengthString(saveName);
            this.LengthMillis = AudioHelper.GetAudioLength(saveName);
        }

        public static AudioType GetShapeAudioType(Shape shape)
        {
            AudioType audioType = AudioType.Unrecognized;
            if (!Enum.TryParse<AudioType>(shape.Tags[AudioTypeTagName], out audioType))
            {
                // Maintain backwards compatibility with old audio shapes
                // which still use sampling rate to store
                switch (shape.MediaFormat.AudioSamplingRate)
                {
                    case GeneratedSamplingRate:
                        audioType = AudioType.Auto;
                        break;
                    case RecordedSamplingRate:
                        audioType = AudioType.Record;
                        break;
                    default:
                        audioType = AudioType.Unrecognized;
                        break;
                }
            }
            return audioType;
        }

        // before we embed we need to check if we have any old shape on the slide. If
        // we have, we need to delete it AFTER the new shape is inserted to preserve
        // the original timeline.
        public void EmbedOnSlide(PowerPointSlide slide, int clickNumber)
        {
            bool isOnClick = clickNumber > 0;
            string shapeName = Name;

            if (slide != null)
            {
                // embed new shape using two-turn method. In the first turn, embed the shape, name it to
                // something special to distinguish from the old shape; in the second turn, delete the
                // old shape using timeline invariant deletion, and rename the new shape to the correct
                // name.
                try
                {
                    Shape audioShape = AudioHelper.InsertAudioFileOnSlide(slide, SaveName);
                    audioShape.Name = "#";
                    slide.RemoveAnimationsForShape(audioShape);

                    if (isOnClick)
                    {
                        slide.SetShapeAsClickTriggered(audioShape, clickNumber, MsoAnimEffect.msoAnimEffectMediaPlay);
                    }
                    else
                    {
                        slide.SetAudioAsAutoplay(audioShape);
                    }

                    // delete old shape
                    slide.DeleteShapesWithPrefixTimelineInvariant(Name);

                    audioShape.Name = shapeName;
                    
                    // tag item with type
                    audioShape.Tags.Add(AudioTypeTagName, Type.ToString());
                }
                catch (COMException)
                {
                    // Adding the file failed for one reason or another - probably cancelled by the user.
                }
            }
            else
            {
                WPFMessageBox.Show(TextCollection.CommonText.ErrorSlideSelectionTitle);
            }
        }
    }
}
