using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WordToPDF;

namespace ConvertFromWordToPDF
{
    class Program
    {
        static bool validMenuOptionSelected = false;
        static bool validFileInput = false;

        static Dictionary<string, string> Directories = new Dictionary<string, string>();

        static void Main(string[] args)
        {
            PrintIntro();
            PrintMenu();

            int selectedMenuOption = 0;
            SelectMenuOption(ref selectedMenuOption);

            if (selectedMenuOption == 1)
            {
                PrintSingleConvert();
                GetInputForSingleConvert();
                ConvertMany();
            }
            else if (selectedMenuOption == 2)
            {
                Console.WriteLine("Multiple Convert");
            }
            else if (selectedMenuOption == 3)
            {
                Console.WriteLine("Directory Convert");
            }
            else
            {
                Console.WriteLine("Exiting...");
                Environment.Exit(0);
            }
        }

        static void PrintIntro()
        {
            Console.WriteLine(new string('=', 57));
            Console.WriteLine("A simple console application to convert word files to pdf");
            Console.WriteLine("\thttps://github.com/mohamed-azhar");
            Console.WriteLine(new string('=', 57));
        }

        static void PrintMenu()
        {
            Console.WriteLine("\t\tSelect an Option");
            Console.WriteLine("\t\t\t1 - Single Convert");
            Console.WriteLine("\t\t\t2 - Multiple Convert");
            Console.WriteLine("\t\t\t3 - Directory Convert - Coming Soon");
            Console.WriteLine("\t\t\t4 - Quit");
        }

        static void PrintSingleConvert()
        {
            Console.WriteLine("\nSingle Convert");
            Console.WriteLine(new string('=', 57));
        }

        static void GetInputForSingleConvert()
        {
            string docExtention = ".doc";
            string docxExtention = ".docx";
            string filePath = string.Empty;
            string targetLocation = string.Empty;

            do
            {
                Console.Write("Fully qualified path of the file(including the file name): ");
                filePath = Console.ReadLine();

                if (File.Exists($"{filePath}{docExtention}"))
                {
                    validFileInput = true;
                    filePath += docExtention;
                }
                else if(File.Exists($"{filePath}{docxExtention}"))
                {
                    validFileInput = true;
                    filePath += docxExtention;
                }
                else
                {
                    Console.WriteLine("File doesnt exist.");
                    validFileInput = false;
                }

                if (validFileInput)
                {
                    Console.Write("Target location for the converted file (leave blank if for the current directory): ");
                    targetLocation = Console.ReadLine();

                    if (string.IsNullOrWhiteSpace(targetLocation))
                    {
                        targetLocation = GetDirectoryFromFilePath(filePath);
                    }

                    if (!Directory.Exists(targetLocation))
                    {
                        Console.WriteLine("Path doesnt exist so will be created.");
                        try
                        {
                            Directory.CreateDirectory(targetLocation);
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("Something went wrong while creating the specified path.");
                            Environment.Exit(0);
                        }
                    }
                }

            } while (!validFileInput);


            Directories.Add(filePath, string.IsNullOrWhiteSpace(targetLocation) ? filePath : targetLocation);
        }

        static void SelectMenuOption(ref int selectedMenuOption)
        {
            do
            {
                Console.Write("select an option: ");
                int.TryParse(Console.ReadLine(), out int option);

                if (option < 5 && option > 0)
                {
                    validMenuOptionSelected = true;
                    selectedMenuOption = option;
                }
                else
                {
                    Console.WriteLine("Invalid option\n");
                    validMenuOptionSelected = false;
                }

            } while (!validMenuOptionSelected);
        }

        static void ConvertMany()
        {
            if (Directories?.Any() ?? false)
            {
                foreach (var (key, value) in Directories)
                {
                    ConvertSingle(key, value);
                }
            }
        }

        static void ConvertSingle(string source, string destination)
        {
            if (!string.IsNullOrWhiteSpace(source) && !string.IsNullOrWhiteSpace(destination))
            {
                string pdfExtension = ".pdf";
                var extension = Path.GetExtension(source);

                if (IsValidWordFile(extension))
                {
                    var fileName = GetFileName(source);

                    var converter = new Word2Pdf()
                    {
                        InputLocation = source,
                        OutputLocation = Path.Combine(destination, $"{fileName}{pdfExtension}")
                    };

                    if (source == destination)
                    {
                        var sourceExtension = Path.GetExtension(source);
                        converter.OutputLocation = source.Replace(sourceExtension, pdfExtension);
                    }

                    Console.WriteLine("\nConverting File: " + fileName);

                    try
                    {
                        var result = converter.Word2PdfCOnversion();
                        Console.WriteLine("Done...\n");
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("Something went wrong.\n");
                    }

                }
            }
        }

        static string[] GetFilesInDirectory(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new ArgumentNullException(nameof(directory));
            }

            var directoryInformation = new DirectoryInfo(directory);
            return directoryInformation.EnumerateFiles(null, SearchOption.TopDirectoryOnly).Select(x => x.FullName).ToArray();
        }

        static bool IsValidWordFile(string extension) => extension == ".doc" || extension == ".docx";
        static string GetFileName(string path) => path.Split("\\").LastOrDefault().Split(".").FirstOrDefault();

        static string GetDirectoryFromFilePath(string filePath)
        {
            var splitted = filePath.Split("\\").Reverse().Skip(1).ToArray();
            return string.Join("\\", splitted.Reverse());
        }
    }
}
