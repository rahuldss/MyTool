using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using  System.Runtime.InteropServices;

namespace CSharpCodes
{
    
    public class Wave
    {
        public static void MakeSound()
        {
            FileStream stream = new FileStream("test.wav", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(stream);
            int RIFF = 0x46464952;
            int WAVE = 0x45564157;
            int formatChunkSize = 16;       //Can't be changed! Value is 16
            int headerSize = 8;             //8
            int format = 0x20746D66;        //0x20746D66
            short formatType = 1;           //1
            short tracks = 2;               //1         //Low - Mid - High
            int samplesPerSecond = 44100;   //44100     //Duration
            short bitsPerSample = 16;       //16        //Duration
            short frameSize = (short)(tracks * ((bitsPerSample + 7) / 8));
            int bytesPerSecond = samplesPerSecond * frameSize;
            int waveSize = 4;               //4
            int data = 0x61746164;          //0x61746164
            int samples = 88200 * 4;        //88200 * 4 //Duration
            int dataChunkSize = samples * frameSize;
            int fileSize = waveSize + headerSize + formatChunkSize + headerSize + dataChunkSize;
            writer.Write(RIFF);
            writer.Write(fileSize);
            writer.Write(WAVE);
            writer.Write(format);
            writer.Write(formatChunkSize);
            writer.Write(formatType);
            writer.Write(tracks);
            writer.Write(samplesPerSecond);
            writer.Write(bytesPerSecond);
            writer.Write(frameSize);
            writer.Write(bitsPerSample);
            writer.Write(data);
            writer.Write(dataChunkSize);
            double aNatural = 220.0;
            double ampl = 200000;            //10000
            double perfect = 2;             //1.5
            double concert = 1.498307077;
            double freq = aNatural * perfect;
            for (int i = 0; i < samples / 4; i++)
            {
                double t = (double)i / (double)samplesPerSecond;
                short s = (short)(ampl * (Math.Sin(t * freq * 2.0 * Math.PI)));
                writer.Write(s);
            }
            freq = aNatural * concert;
            for (int i = 0; i < samples / 4; i++)
            {
                double t = (double)i / (double)samplesPerSecond;
                short s = (short)(ampl * (Math.Sin(t * freq * 2.0 * Math.PI)));
                writer.Write(s);
            }
            for (int i = 0; i < samples / 4; i++)
            {
                double t = (double)i / (double)samplesPerSecond;
                short s = (short)(ampl * (Math.Sin(t * freq * 2.0 * Math.PI) + Math.Sin(t * freq * perfect * 2.0 * Math.PI)));
                writer.Write(s);
            }
            for (int i = 0; i < samples / 4; i++)
            {
                double t = (double)i / (double)samplesPerSecond;
                short s = (short)(ampl * (Math.Sin(t * freq * 2.0 * Math.PI) + Math.Sin(t * freq * concert * 2.0 * Math.PI)));
                writer.Write(s);
            }
            writer.Close();
            stream.Close();
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool Beep(uint dwFreq, uint dwDuration);
        public static void Beep()
        {
            Console.WriteLine("Testing PC speaker...");
            for (uint i = 100; i <= 20000; i++)
            {
                Beep(i, 5);
            }
            Console.WriteLine("Testing complete.");
        }


    }
}