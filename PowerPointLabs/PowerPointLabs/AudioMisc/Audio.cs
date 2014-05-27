using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

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
        public const int RecordedBitRate = 8;

        public string Name { get; set; }
        public int MatchSciptID { get; set; }
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
            MatchSciptID = matchScriptID;
            SaveName = saveName;
            Length = AudioHelper.GetAudioLengthString(saveName);
            LengthMillis = AudioHelper.GetAudioLength(saveName);
            Type = AudioHelper.GetAudioType(saveName);
        }

        public void EmbedOnSlide(PowerPointSlide slide)
        {
            var shapeName = Name;
            var isOnClick = SaveName.Contains("OnClick");

            // init click number
            var temp = Name.Split(new[] { ' ' });
            var clickNumber = Int32.Parse(temp[2]);

            if (slide != null)
            {
                // delete old shape
                slide.DeleteShapesWithPrefix(shapeName);

                // embed new shape
                try
                {
                    var audioShape = AudioHelper.InsertAudioFileOnSlide(slide, SaveName);
                    slide.RemoveAnimationsForShape(audioShape);
                    audioShape.Name = shapeName;

                    if (isOnClick)
                    {
                        slide.SetShapeAsClickTriggered(audioShape, clickNumber, MsoAnimEffect.msoAnimEffectMediaPlay);
                    }
                    else
                    {
                        slide.SetAudioAsAutoplay(audioShape);
                    }
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
