using Colorful;
using System;
using System.Drawing;
using Console = System.Console;
using ColorfulConsole = Colorful.Console;

namespace ConvertFromWordToPDF.Helpers
{
    public static class DisplayHelpers
    {
        public static void PrintIntro()
        {
            //Console.WriteLine(new string('=', 57));
            //Console.WriteLine("A simple console application to convert word files to pdf");
            //Console.WriteLine("\thttps://github.com/mohamed-azhar");
            //Console.WriteLine(new string('=', 57));
            ColorfulConsole.WriteAscii("WORD 2 PDF", Color.FromArgb(244, 212, 255));
        }

        public static void PrintMenu()
        {
            string select = "\t\t{0}";
            string single = "\t\t\t1 - {0}";
            string directory = "\t\t\t2 - {0}";
            string quit = "\t\t\t3 - {0}";

            ColorfulConsole.WriteLineFormatted(select, Color.Gray, new Formatter("Select an Option", Color.LawnGreen));
            ColorfulConsole.WriteLineFormatted(single, Color.Gray, new Formatter("Single Convert", Color.LawnGreen));
            ColorfulConsole.WriteLineFormatted(directory, Color.Gray, new Formatter("Directory Convert", Color.LawnGreen));
            ColorfulConsole.WriteLineFormatted(quit, Color.Gray, new Formatter("Quit", Color.LawnGreen));
        }

        public static void PrintSingleConvert()
        {
            Console.WriteLine("\nSingle Convert");
            Console.WriteLine(new string('=', 57));
        }

        public static void PrintDirectoryConvert()
        {
            Console.WriteLine("\nDirectory Convert");
            Console.WriteLine(new string('=', 57));
        }
    }
}
