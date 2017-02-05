using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;

namespace PowerPointLabs.AudioMisc
{
    internal class Audio
    {
        public enum AudioType
        {
            Record,
            Auto
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

        /// <summary>
        /// Default constructor.
        /// </summary>
        public Audio() {}

        public Audio(string name, string saveName, int matchScriptID)
        {
            Name = name;
            MatchScriptID = matchScriptID;
            SaveName = saveName;
            LengthMillis = AudioHelper.GetAudioLength(saveName);
            Length = AudioHelper.ConvertMillisToTime(LengthMillis);
            Type = AudioHelper.GetAudioType(saveName);
        }

        // before we embed we need to check if we have any old shape on the slide. If
        // we have, we need to delete it AFTER the new shape is inserted to preserve
        // the original timeline.
        public void EmbedOnSlide(PowerPointSlide slide, int clickNumber)
        {
            var isOnClick = clickNumber > 0;
            var shapeName = Name;

            if (slide != null)
            {
                // embed new shape using two-turn method. In the first turn, embed the shape, name it to
                // something special to distinguish from the old shape; in the second turn, delete the
                // old shape using timeline invariant deletion, and rename the new shape to the correct
                // name.
                try
                {
                    var audioShape = AudioHelper.InsertAudioFileOnSlide(slide, SaveName);
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
                }
                catch (COMException)
                {
                    // Adding the file failed for one reason or another - probably cancelled by the user.
                }
            }
            else
            {
                MessageBox.Show("Slide selection error");
            }
        }
    }
}
