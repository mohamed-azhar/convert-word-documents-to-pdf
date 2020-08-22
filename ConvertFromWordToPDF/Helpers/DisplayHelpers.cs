using System;

namespace ConvertFromWordToPDF.Helpers
{
    public static class DisplayHelpers
    {
        public static void PrintIntro()
        {
            Console.WriteLine(new string('=', 57));
            Console.WriteLine("A simple console application to convert word files to pdf");
            Console.WriteLine("\thttps://github.com/mohamed-azhar");
            Console.WriteLine(new string('=', 57));
        }

        public static void PrintMenu()
        {
            Console.WriteLine("\t\tSelect an Option");
            Console.WriteLine("\t\t\t1 - Single Convert");
            Console.WriteLine("\t\t\t2 - Multiple Convert");
            Console.WriteLine("\t\t\t3 - Directory Convert - Coming Soon");
            Console.WriteLine("\t\t\t4 - Quit");
        }

        public static void PrintSingleConvert()
        {
            Console.WriteLine("\nSingle Convert");
            Console.WriteLine(new string('=', 57));
        }
    }
}
