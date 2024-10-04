using System;
using System.Configuration;
using System.IO;
using GroupDocs.Conversion;
using GroupDocs.Conversion.Options.Convert;

namespace ODTtoPDF_Converter_using_CSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Read source and destination paths from the app.config file
                string sourceDirectory = ConfigurationManager.AppSettings["SourceDirectory"];
                string destinationDirectory = ConfigurationManager.AppSettings["DestinationDirectory"];

                // Ensure the destination directory exists
                if (!Directory.Exists(destinationDirectory))
                {
                    Directory.CreateDirectory(destinationDirectory);
                }

                // Get all ODT files from the source directory
                string[] odtFiles = Directory.GetFiles(sourceDirectory, "*.odt");

                if (odtFiles.Length == 0)
                {
                    Console.WriteLine("No ODT files found in the source directory.");
                    return;
                }

                foreach (string odtFilePath in odtFiles)
                {
                    try
                    {
                        // Get the file name without extension
                        string fileName = Path.GetFileNameWithoutExtension(odtFilePath);

                        // Set the output PDF file path
                        string outputFilePath = Path.Combine(destinationDirectory, fileName + ".pdf");

                        // Load the ODT document for conversion
                        using (Converter converter = new Converter(odtFilePath))
                        {
                            PdfConvertOptions options = new PdfConvertOptions();
                            // Convert ODT to PDF
                            converter.Convert(outputFilePath, options);
                        }

                        Console.WriteLine($"Successfully converted '{odtFilePath}' to PDF.");

                        // Optionally, delete the source ODT file after conversion
                        File.Delete(odtFilePath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error converting file {odtFilePath}: {ex.Message}");
                    }
                }

                Console.WriteLine("Conversion process completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}