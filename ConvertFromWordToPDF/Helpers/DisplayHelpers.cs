using Colorful;
using ConvertFromWordToPDF.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.Linq;
using ColorfulConsole = Colorful.Console;

namespace ConvertFromWordToPDF.Helpers
{
    public static class DisplayHelpers
    {
        public static void PrintIntro()
        {
            ColorfulConsole.WriteAscii("WORD 2 PDF", Color.FromArgb(32, 106, 93));
        }

        public static void PrintMenu()
        {
            var color = Color.DarkGreen;
            var content = new Dictionary<string, Formatter>()
            {
                { "\t\t{0}", new Formatter("Select an Option", color)},
                { "\t\t     1 - {0}",  new Formatter("Single File Convert", color)},
                { "\t\t     2 - {0}",  new Formatter("Full Directory Convert", color)},
                { "\t\t     3 - {0}",  new Formatter("Quit", color)},
            };

            PrintFormatted(content, Color.White, ConsoleWriteMethod.WriteLine);
        }

        public static void PrintSingleConvert()
        {
            var color = Color.DarkGreen;
            var content = new Dictionary<string, Formatter>()
            {
                { "\n{0}", new Formatter("Single File Convert", color)},
                { "{0}",  new Formatter($"{new string('=', 57)}", color)},
            };

            PrintFormatted(content, Color.White, ConsoleWriteMethod.WriteLine);
        }

        public static void PrintDirectoryConvert()
        {
            var color = Color.DarkGreen;
            var content = new Dictionary<string, Formatter>()
            {
                { "\n{0}", new Formatter("Full Directory Convert", color)},
                { "{0}",  new Formatter($"{new string('=', 57)}", color)},
            };

            PrintFormatted(content, Color.White, ConsoleWriteMethod.WriteLine);
        }

        public static void DisplayValidFileNames(string[] fileNams)
        {
            if (fileNams?.Any() ?? false)
            {
                for (int i = 0; i < fileNams.Length; i++)
                {
                    var current = fileNams[i];
                    Print($"\t{i+1} - {FileHelpers.GetFileName(current)}", Color.Yellow, ConsoleWriteMethod.WriteLine);
                }
                System.Console.WriteLine();
            }
        }

        public static void PrintFormatted(Dictionary<string, Formatter> textAndFormats, Color fallbackColor, ConsoleWriteMethod writeMethod)
        {
            if (textAndFormats?.Any() ?? false)
            {
                foreach (var (text, format) in textAndFormats)
                {
                    switch (writeMethod)
                    {
                        case ConsoleWriteMethod.Write:
                            ColorfulConsole.WriteFormatted(text, fallbackColor, format);
                            break;
                        case ConsoleWriteMethod.WriteLine:
                        default:
                            ColorfulConsole.WriteLineFormatted(text, fallbackColor, format);
                            break;
                    }
                }
            }
        }

        public static void Print(string text, Color color, ConsoleWriteMethod writeMethod)
        {
            if (!string.IsNullOrWhiteSpace(text))
            {
                switch (writeMethod)
                {
                    case ConsoleWriteMethod.Write:
                        ColorfulConsole.Write(text, color);
                        break;
                    case ConsoleWriteMethod.WriteLine:
                    default:
                        ColorfulConsole.WriteLine(text, color);
                        break;
                }
            }
        }
    }
}
