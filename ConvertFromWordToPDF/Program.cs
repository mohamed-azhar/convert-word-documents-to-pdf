using Colorful;
using ConvertFromWordToPDF.Enums;
using ConvertFromWordToPDF.Helpers;
using System;
using System.Collections.Generic;
using System.Drawing;
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
        static IDictionary<string, string> Directories = new Dictionary<string, string>();

        static void Main(string[] args)
        {
            DisplayHelpers.PrintIntro();
            DisplayHelpers.PrintMenu();

            while (true)
            {
                int selectedMenuOption = SelectMenuOption();

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
                                DisplayHelpers.Print($"\nFound {validFileNames.Length} files which can be convert to PDF.", Color.Yellow, ConsoleWriteMethod.WriteLine);

                                //display all the valid file names
                                DisplayHelpers.DisplayValidFileNames(validFileNames);

                                //get output path
                                var outputPath = GetOutputDirectory();

                                //set output path to be default source path if blank
                                if (string.IsNullOrWhiteSpace(outputPath))
                                {
                                    outputPath = GetOutputDirectoryIfEmptyInput(validFileNames?.FirstOrDefault());
                                }

                                //prepare the directories dictionary with the valid filenames and their output path
                                foreach (var fileName in validFileNames)
                                {
                                    Directories.Add(fileName, outputPath);
                                }

                                //convert the input files
                                ConvertMany();
                                DisplayHelpers.Print($"\nSuccessfully converted {validFileNames.Length} Files.", Color.DarkGreen, ConsoleWriteMethod.WriteLine);
                            }
                            else
                            {
                                DisplayHelpers.Print("No valid word files found in the provided directory.", Color.DarkRed, ConsoleWriteMethod.WriteLine);
                            }

                            //clear the directories
                            ClearDirectories();
                            break;
                        }
                    case (int)MenuOption.Quit:
                        DisplayHelpers.Print("Exiting...\n", Color.Gray, ConsoleWriteMethod.WriteLine);
                        Environment.Exit(0);
                        break;
                    default:
                        DisplayHelpers.Print("Exiting...\n", Color.Gray, ConsoleWriteMethod.WriteLine);
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
                DisplayHelpers.Print("Fully qualified path of the file(including the file name): ", Color.Gray, ConsoleWriteMethod.Write);
                filePath = System.Console.ReadLine();

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
                    DisplayHelpers.Print("File doesnt exist.\n", Color.DarkRed, ConsoleWriteMethod.WriteLine);
                    validFileInput = false;
                }

                if (validFileInput)
                {
                    DisplayHelpers.Print("Target location for the converted file (leave blank if for the current directory): ", Color.Gray, ConsoleWriteMethod.Write);
                    targetLocation = System.Console.ReadLine();

                    if (string.IsNullOrWhiteSpace(targetLocation))
                    {
                        targetLocation = FileHelpers.GetDirectoryFromFilePath(filePath);
                    }

                    if (!Directory.Exists(targetLocation))
                    {
                        DisplayHelpers.Print("Path doesnt exist so will be created.", Color.Yellow, ConsoleWriteMethod.WriteLine);
                        try
                        {
                            Directory.CreateDirectory(targetLocation);
                        }
                        catch (Exception)
                        {
                            DisplayHelpers.Print("Something went wrong while creating the specified path.", Color.DarkRed, ConsoleWriteMethod.WriteLine);
                            Environment.Exit(0);
                        }
                    }
                }

            } while (!validFileInput);


            Directories.Add(filePath, string.IsNullOrWhiteSpace(targetLocation) ? filePath : targetLocation);
        }

        static string GetInputForDirectoryConvert()
        {
            string path = string.Empty;
            do
            {
                DisplayHelpers.Print("Enter directory path: ", Color.Gray, ConsoleWriteMethod.Write);
                path = System.Console.ReadLine();

                if (Directory.Exists(path))
                {
                    validDirectoryInput = true;
                    return path;
                }
                else
                {
                    validDirectoryInput = false;
                    DisplayHelpers.Print("Invalid directory.\n", Color.DarkRed, ConsoleWriteMethod.WriteLine);
                }
            } while (!validDirectoryInput);
            return path;
        }

        static int SelectMenuOption()
        {
            var count = (int)Enum.GetValues(typeof(MenuOption)).Cast<MenuOption>().Max();
            int selectedOption = 0;
            do
            {
                DisplayHelpers.Print("\nselect an option: ", Color.Gray, ConsoleWriteMethod.Write);
                int.TryParse(System.Console.ReadLine(), out int option);

                if (option < (count + 1) && option > 0)
                {
                    validMenuOptionSelected = true;
                    selectedOption = option;
                }
                else
                {
                    DisplayHelpers.Print("Invalid option.", Color.DarkRed, ConsoleWriteMethod.WriteLine);
                    validMenuOptionSelected = false;
                }

            } while (!validMenuOptionSelected);

            return selectedOption;
        }

        static string GetOutputDirectory()
        {
            string path = string.Empty;
            do
            {
                DisplayHelpers.Print("Enter output path (leave blank if for the current directory): ", Color.Gray, ConsoleWriteMethod.Write);
                path = System.Console.ReadLine();

                if (Directory.Exists(path))
                {
                    validFileOutput = true;
                    return path;
                }
                else
                {
                    try
                    {
                        if (string.IsNullOrWhiteSpace(path))
                        {
                            validFileOutput = true;
                        }
                        else
                        {
                            DisplayHelpers.Print("Directory doesnt exist so will be created.", Color.Yellow, ConsoleWriteMethod.WriteLine);
                            Directory.CreateDirectory(path);
                            validFileOutput = true;
                        }
                        return path;
                    }
                    catch (Exception)
                    {
                        validFileOutput = false;
                        DisplayHelpers.Print("Invalid directory.\n", Color.DarkRed, ConsoleWriteMethod.WriteLine);
                    }
                }
            } while (!validFileOutput);
            return path;
        }

        static string GetOutputDirectoryIfEmptyInput(string filePath)
        {
            return FileHelpers.GetDirectoryFromFilePath(filePath);
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

                    DisplayHelpers.Print("\nConverting File: " + fileName, Color.Yellow, ConsoleWriteMethod.WriteLine);

                    try
                    {
                        var result = converter.Word2PdfCOnversion();
                        DisplayHelpers.Print("Done..", Color.DarkGreen, ConsoleWriteMethod.WriteLine);
                    }
                    catch (Exception ex)
                    {
                        DisplayHelpers.Print($"Something went wrong. Message: {ex.Message}.\n", Color.DarkRed, ConsoleWriteMethod.WriteLine);
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
