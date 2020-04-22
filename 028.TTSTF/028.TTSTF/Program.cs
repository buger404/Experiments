using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Speech.AudioFormat;
using System.Speech.Synthesis;
using System.Text;
using System.Threading.Tasks;

namespace _028.TTSTF
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("TTS to File, Output Path: 'D:\\TTSTF\'\n");
                if (!Directory.Exists("D:\\TTSTF")) Directory.CreateDirectory("D:\\TTSTF");
                SpeechSynthesizer reader = new SpeechSynthesizer();
                reader.Volume = 100;
                Console.Write("Speaking rate(-10~+10,default:0):");
                reader.Rate = int.Parse(Console.ReadLine());
            redo:
                Console.ForegroundColor = ConsoleColor.White;
                Console.Write("Input:");
                string word = Console.ReadLine();
                reader.SetOutputToWaveFile($"D:\\TTSTF\\{word}.wav", new SpeechAudioFormatInfo(32000, AudioBitsPerSample.Sixteen, AudioChannel.Mono));
                //reader.Rate = -2 + new Random(Guid.NewGuid().GetHashCode()).Next(0, 4);
                PromptBuilder builder = new PromptBuilder();
                builder.AppendText(word);
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine("Start speaking...");
                reader.Speak(builder);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Outputed!\n");
                goto redo;
            }
            catch(Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("An error when speaking: " + e.Message);
                Main(args);
            }
        }
    }
}
