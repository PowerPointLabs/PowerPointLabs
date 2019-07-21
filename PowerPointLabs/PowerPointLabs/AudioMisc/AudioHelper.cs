using System;
using System.Text;

using Microsoft.Office.Core;

using PowerPointLabs.Models;

using PPExtraEventHelper;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.AudioMisc
{
    internal class AudioHelper
    {
#pragma warning disable 0618
        #region MCI Constants
        public const int MM_MCINOTIFY = 0x03B9;
        public const int MCI_NOTIFY_SUCCESS = 0x01;
        public const int MCI_NOTIFY_ABORTED = 0x04;
        public const int MCI_NOTIFY_FAILURE = 0x08;

        private const int MCI_RET_INFO_BUF_LEN = 128;

        private static StringBuilder mciRetInfo;
        # endregion

        /// <summary>
        /// This function will convert a time in milli-second to HH:MM:SS:MMS
        /// </summary>
        /// <param name="millis">Time in millis.</param>
        /// <returns>A string in HH:MM:SS:MMS format.</returns>
        public static string ConvertMillisToTime(long millis)
        {
            int ms, s, m, h;

            ms = (int)millis % 1000;
            millis /= 1000;

            s = (int)(millis % 60);
            millis /= 60;

            m = (int)(millis % 60);
            millis /= 60;

            h = (int)(millis % 60);
            millis /= 60;

            return System.String.Format("{0:D2}:{1:D2}:{2:D2}", h, m, s);
        }

        public static string ConvertMillisToTime(int millis)
        {
            return ConvertMillisToTime((long)millis);
        }

        public static void OpenNewAudio()
        {
            Native.mciSendString("open new type waveaudio alias sound", null, 0, IntPtr.Zero);
        }

        public static void OpenAudio(string name)
        {
            int x = Native.mciSendString("open \"" + name + "\" alias sound", null, 0, IntPtr.Zero);
        }

        public static void CloseAudio()
        {
            Native.mciSendString("close sound", null, 0, IntPtr.Zero);
        }

        public static int GetAudioLength()
        {
            mciRetInfo = new StringBuilder(MCI_RET_INFO_BUF_LEN);
            int x = Native.mciSendString("status sound length", mciRetInfo, MCI_RET_INFO_BUF_LEN, IntPtr.Zero);
            return Int32.Parse(mciRetInfo.ToString());
        }

        public static int GetAudioLength(string name)
        {
            OpenAudio(name);
            int length = GetAudioLength();
            CloseAudio();

            return length;
        }

        public static string GetAudioLengthString()
        {
            int length = GetAudioLength();
            return ConvertMillisToTime(length);
        }

        public static string GetAudioLengthString(string name)
        {
            OpenAudio(name);
            string length = GetAudioLengthString();
            CloseAudio();

            return length;
        }

        public static int GetAudioCurrentPosition()
        {
            mciRetInfo = new StringBuilder(MCI_RET_INFO_BUF_LEN);
            Native.mciSendString("status sound position", mciRetInfo, MCI_RET_INFO_BUF_LEN, IntPtr.Zero);
            return Int32.Parse(mciRetInfo.ToString());
        }

        public static Audio.AudioType GetAudioType(string name)
        {
            if (name.Contains("rec"))
            {
                return Audio.AudioType.Record;
            }

            return Audio.AudioType.Auto;
        }

        /// <summary>
        /// Dump current sound into an Audio object.
        /// </summary>
        /// <param name="name">The name that display on panes.</param>
        /// <param name="saveName">The intended save name.</param>
        /// <param name="matchScriptID">The corresponding script id for the current track.</param>
        /// <returns>An Audio object with all fields set up.</returns>
        public static Audio DumpAudio(string name, string saveName, int length, int matchScriptID)
        {
            // get length from file if length is not specified
            if (length == -1)
            {
                length = GetAudioLength(saveName);  
            }

            Audio audio = new Audio
                            {
                                Name = name,
                                SaveName = saveName,
                                LengthMillis = length,
                                Length = ConvertMillisToTime(length),
                                MatchScriptID = matchScriptID,
                                Type = GetAudioType(saveName)
                            };
            
            return audio;
        }
    }
}
