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

        public void EmbedOnSlide(PowerPointSlide slide)
        {
            // to distinguish the new shape with the old shape
            var shapeName = Name;
            var isOnClick = SaveName.Contains("OnClick");

            // init click number
            var temp = Name.Split(new[] { ' ' });
            var clickNumber = Int32.Parse(temp[2]);

            if (slide != null)
            {
                var sequence = slide.TimeLine.MainSequence;
                var nextClickEffect = sequence.FindFirstAnimationForClick(clickNumber + 1);

                // embed new shape
                try
                {
                    var audioShape = AudioHelper.InsertAudioFileOnSlide(slide, SaveName);
                    //slide.RemoveAnimationsForShape(audioShape);
                    
                    // give a temp name first so that we are able to delete the old shape
                    // with prefix
                    audioShape.Name = "#";

                    if (isOnClick)
                    {
                        //slide.SetShapeAsClickTriggered(audioShape, clickNumber, MsoAnimEffect.msoAnimEffectMediaPlay);

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
