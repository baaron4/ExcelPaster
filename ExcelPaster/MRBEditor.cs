using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPaster
{
    public class MRBEditor
    {

        public void FindReplaceValue(string inputFilePath, string outputFilePath, byte searchValue, byte replaceValue)
        {
            byte[] fileContent = File.ReadAllBytes(inputFilePath);

            for (int i = 0; i < fileContent.Length; i++)
            {
                byte currentValue = (byte)fileContent[i];
                if (currentValue == searchValue)
                {
                    fileContent[i] = (byte)replaceValue;
                }
            }

            File.WriteAllBytes(outputFilePath, fileContent);
        }

        public void ProcessFiles(string inputDirectoryPath, string outputDirectoryPath, byte searchValue, byte replaceValue)
        {
            if (!Directory.Exists(outputDirectoryPath))
            {
                Directory.CreateDirectory(outputDirectoryPath);
            }

            string[] fileNames = Directory.GetFiles(inputDirectoryPath, "*.mrb");

            foreach (string inputFilePath in fileNames)
            {
                string fileName = Path.GetFileName(inputFilePath);
                string outputFilePath = Path.Combine(outputDirectoryPath, fileName);
                FindReplaceValue(inputFilePath, outputFilePath, searchValue, replaceValue);
                Console.WriteLine($"Replaced values in {fileName}. Saved as {outputFilePath}");
            }
        }
    }
}
