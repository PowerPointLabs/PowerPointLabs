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

        // before we embed, we need to check if we have any old shape on the slide if
        // we have, we need to delete it AFTER the new shape is inserted to preserve
        // the original timeline.
        // However, for the inserted shape animation, we need to see if the animation
        // is the first animation for the click. If there are some other click events
        // on that click, the animation should be made as (OnPrev); else it should be
        // made as (OnPageClick).
        public void EmbedOnSlide(PowerPointSlide slide, int clickNumber)
        {
            // to distinguish the new shape with the old shape
            var shapeName = Name;
            var isOnClick = clickNumber > 0;

            if (slide != null)
            {
                var sequence = slide.TimeLine.MainSequence;
                var nextClickEffect = sequence.FindFirstAnimationForClick(clickNumber + 1);

                // embed new shape
                try
                {
                    var audioShape = AudioHelper.InsertAudioFileOnSlide(slide, SaveName);
                    
                    // give a temp name first so that we are able to delete the old shape
                    // with prefix
                    audioShape.Name = "#";

                    if (isOnClick)
                    {
                        if (nextClickEffect != null)
                        {
                            var newAnimation = sequence.AddEffect(audioShape, MsoAnimEffect.msoAnimEffectMediaPlay,
                                                              MsoAnimateByLevel.msoAnimateLevelNone,
                                                              MsoAnimTriggerType.msoAnimTriggerOnPageClick, clickNumber);
                            newAnimation.MoveBefore(nextClickEffect);
                        }
                        else
                        {
                            sequence.AddEffect(audioShape, MsoAnimEffect.msoAnimEffectMediaPlay);
                        }
                    }
                    else
                    {
                        slide.SetAudioAsAutoplay(audioShape);
                    }
                    
                    // delete old shape
                    slide.DeleteShapesWithPrefix(shapeName);

                    // rename the new shape
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
