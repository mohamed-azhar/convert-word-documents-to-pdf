using ConvertFromWordToPDF.Enums;
using ConvertFromWordToPDF.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WordToPDF;

namespace ConvertFromWordToPDF
{
    class Program
    {
        static bool validFileInput = false;
        static bool validFileOutput = false;
        static bool validDirectoryInput = false;
        static bool validMenuOptionSelected = false;

        /// <summary>
        /// Directionary containing source filename and output path name
        /// </summary>
        static Dictionary<string, string> Directories = new Dictionary<string, string>();

        static void Main(string[] args)
        {
            DisplayHelpers.PrintIntro();
            DisplayHelpers.PrintMenu();

            while (true)
            {
                int selectedMenuOption = 0;
                SelectMenuOption(ref selectedMenuOption);

                switch (selectedMenuOption)
                {
                    case (int)MenuOption.SingleConvert:
                        {
                            //print single convert text
                            DisplayHelpers.PrintSingleConvert();

                            //get input for single convert
                            GetInputForSingleConvert();

                            //convert the input file
                            ConvertMany();

                            //clear the directories
                            ClearDirectories();
                            break;
                        }
                    case (int)MenuOption.DirectoryConvert:
                        {
                            //print directory convert text
                            DisplayHelpers.PrintDirectoryConvert();

                            //get directory which needs to be converted
                            var directoryPath = GetInputForDirectoryConvert();

                            //get all the files from the directory
                            var fileNames = FileHelpers.GetFilesInDirectory(directoryPath);

                            //get valid word files from the files in the given directory
                            var validFileNames = FileHelpers.ExtractValidWordFiles(fileNames);

                            if (validFileNames?.Any() ?? false)
                            {
                                //get output path
                                var outputPath = GetOutputDirectory();

                                //prepare the directories dictionary with the valid filenames and their output path
                                foreach (var fileName in validFileNames)
                                {
                                    Directories.Add(fileName, outputPath);
                                }

                                //convert the input file
                                ConvertMany();
                            }
                            else
                            {
                                Console.WriteLine("No valid word files found in the provided directory.\n");
                            }

                            //clear the directories
                            ClearDirectories();
                            break;
                        }
                    case (int)MenuOption.Quit:
                        Console.WriteLine("Exiting...");
                        Environment.Exit(0);
                        break;
                    default:
                        Console.WriteLine("Exiting...");
                        Environment.Exit(0);
                        break;
                }
            }
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
                else if (File.Exists($"{filePath}{docxExtention}"))
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
                        targetLocation = FileHelpers.GetDirectoryFromFilePath(filePath);
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

        static string GetInputForDirectoryConvert()
        {
            do
            {
                Console.Write("Enter directory path: ");
                string path = Console.ReadLine();

                if (Directory.Exists(path))
                {
                    validDirectoryInput = true;
                    return path;
                }
                else
                {
                    try
                    {
                        Console.WriteLine("Directory doesn't exist so will be created.");
                        Directory.CreateDirectory(path);
                        validDirectoryInput = true;
                        return path;
                    }
                    catch (Exception)
                    {
                        validDirectoryInput = false;
                        Console.WriteLine("Invalid directory\n");
                    }
                }
                return path;
            } while (!validDirectoryInput);
        }

        static void SelectMenuOption(ref int selectedMenuOption)
        {
            var count = (int)Enum.GetValues(typeof(MenuOption)).Cast<MenuOption>().Max();

            do
            {
                Console.Write("select an option: ");
                int.TryParse(Console.ReadLine(), out int option);

                if (option < (count + 1) && option > 0)
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

        static string GetOutputDirectory()
        {
            do
            {
                Console.Write("Enter output path: ");
                string path = Console.ReadLine();

                if (Directory.Exists(path))
                {
                    validFileOutput = true;
                    return path;
                }
                else
                {
                    try
                    {
                        Console.WriteLine("Directory doesn't exist so will be created.");
                        Directory.CreateDirectory(path);
                        validFileOutput = true;

                        return path;
                    }
                    catch (Exception)
                    {
                        validFileOutput = false;
                        Console.WriteLine("Invalid directory\n");
                    }
                }
                return path;
            } while (!validFileOutput);
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

                if (FileHelpers.IsValidWordFile(extension))
                {
                    var fileName = FileHelpers.GetFileName(source);

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
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Something went wrong. Message: {ex.Message}\n");
                    }

                }
            }
        }

        static void ClearDirectories()
        {
            validFileInput = false;
            validMenuOptionSelected = false;
            Directories = new Dictionary<string, string>();
        }
    }
}
