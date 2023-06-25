using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Speech.Synthesis;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DATN
{
    [Serializable]
    public enum CELL_ORDER_TYPE
    {
        INFO,
        WARNING,
        EXCEPTION,
        CRITICAL,
        NONE
    }
    public class SpeechProcessing
    {
        /// <summary>
        ///     Đọc văn bản bằng hàm Cotana. Không cần intenet
        /// </summary>
        /// <param name="text">Văn bản cần đọc</param>
        static public void ReadMeByCotana(string text)
        {
            var synthesizer = new SpeechSynthesizer();
            synthesizer.SetOutputToDefaultAudioDevice();
            var builder = new PromptBuilder();
            builder.StartVoice(new CultureInfo("en-US"));
            builder.AppendText(text);
            builder.EndVoice();
            synthesizer.Speak(builder);
        }
    }
}
