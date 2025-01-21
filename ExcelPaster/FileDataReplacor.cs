using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelPaster
{
    public class FileDataReplacor
    {
        public class PCCURegister
        {
            public UInt16 Index;
            public Byte Array;
            public Byte App;

            public PCCURegister(byte[] raw)
            {
                RawBytesToVariables(raw);
            }

            public PCCURegister(ushort index, byte array, byte app)
            {
                Index = index;
                Array = array;
                App = app;
            }

            public PCCURegister(string raw)
            {
                string[] values = raw.Split('.');
                App = byte.Parse(values[0]);
                Array = byte.Parse(values[1]);
                Index = UInt16.Parse(values[2]);
            }
            public void RawBytesToVariables(byte[] raw)
            {
                App = raw[0];
                Array = raw[1];
                Index = BitConverter.ToUInt16(raw, 2);
            }

            public string ToRegString()
            {
                return App.ToString() + "." + Array.ToString() + "." + Index.ToString();
            }
            public byte[] ToBytes()
            {
                byte[] indexBytes = BitConverter.GetBytes(Index);
                byte[] bytes = { (byte)App, (byte)Array, indexBytes[0], indexBytes[1] };
                
                return bytes;
            }

            public UInt32 ToInt32()
            {
                byte[] indexBytes = BitConverter.GetBytes(Index);
                byte[] bytes = { (byte)App, (byte)Array, indexBytes[0], indexBytes[1] };
                UInt32 integer = BitConverter.ToUInt32(bytes,0);
                return integer;
            }
        }

        public void ReplaceInt16(string inputDirectoryPath, string outputDirectoryPath, byte searchValue, byte replaceValue)
        {
            if (!Directory.Exists(outputDirectoryPath))
            {
                Directory.CreateDirectory(outputDirectoryPath);
            }

            List<string> fileNames = new List<string>();
            fileNames.AddRange(Directory.GetFiles(inputDirectoryPath));

            foreach (string inputFilePath in fileNames)
            {
                string fileName = Path.GetFileName(inputFilePath);
                string outputFilePath = Path.Combine(outputDirectoryPath, fileName);
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
                Console.WriteLine($"Replaced values in {fileName}. Saved as {outputFilePath}");
            }
        }

        public void ReplaceRegister(string inputDirectoryPath, string outputDirectoryPath, string searchValue, string replaceValue)
        {
            if (!Directory.Exists(outputDirectoryPath))
            {
                Directory.CreateDirectory(outputDirectoryPath);
            }

            PCCURegister searchregister = new PCCURegister(searchValue);
            PCCURegister replaceregister = new PCCURegister(replaceValue);
            UInt32 searchInt = searchregister.ToInt32();
            UInt32 replaceInt = replaceregister.ToInt32();
            int numReplacments = 0;

            List<string> fileNames = new List<string>();
            fileNames.AddRange(Directory.GetFiles(inputDirectoryPath));

            foreach (string inputFilePath in fileNames)
            {
                string fileName = Path.GetFileName(inputFilePath);
                string outputFilePath = Path.Combine(outputDirectoryPath, fileName);
                byte[] fileContent = File.ReadAllBytes(inputFilePath);
                int skipByte = 0;
                for (int i = 0; i < fileContent.Length-4; i++)
                {
                    if (skipByte > 0)
                    {
                        skipByte--;
                        continue;
                    }
                    byte[] word = { fileContent[i], fileContent[i + 1], fileContent[i + 2], fileContent[i + 3] };
                    UInt32 currentValue = BitConverter.ToUInt32(word,0);
                    if (currentValue == searchInt)
                    {
                        byte[] replaceByte = BitConverter.GetBytes(replaceInt);
                        fileContent[i] = replaceByte[0];
                        fileContent[i + 1] = replaceByte[1];
                        fileContent[i + 2] = replaceByte[2];
                        fileContent[i + 3] = replaceByte[3];
                        skipByte = 3;
                        numReplacments++;
                    }
                }

                File.WriteAllBytes(outputFilePath, fileContent);
                Console.WriteLine($"Replaced values in {fileName}. Saved as {outputFilePath}");
                
            }
            MessageBox.Show(numReplacments + " Registers replaced");
        }

        public void ReplaceApp(string inputDirectoryPath, string outputDirectoryPath, string searchValue, string replaceValue)
        {
            if (!Directory.Exists(outputDirectoryPath))
            {
                Directory.CreateDirectory(outputDirectoryPath);
            }


            byte searchInt = byte.Parse(searchValue);
            byte replaceInt = byte.Parse(replaceValue);
            int numReplacments = 0;

            List<string> fileNames = new List<string>();
            fileNames.AddRange(Directory.GetFiles(inputDirectoryPath));

            foreach (string inputFilePath in fileNames)
            {
                string fileName = Path.GetFileName(inputFilePath);
                string outputFilePath = Path.Combine(outputDirectoryPath, fileName);
                byte[] fileContent = File.ReadAllBytes(inputFilePath);
                int skipByte = 0;
                for (int i = 0; i < fileContent.Length - 4; i++)
                {
                    if (skipByte > 0)
                    {
                        skipByte--;
                        continue;
                    }
                    byte[] word = { fileContent[i], fileContent[i + 1], fileContent[i + 2], fileContent[i + 3] };
                    byte currentValue = word[0];
                    if (currentValue == searchInt)
                    {
                        fileContent[i] = replaceInt;

                        skipByte = 3;
                        numReplacments++;
                    }
                }

                File.WriteAllBytes(outputFilePath, fileContent);
                Console.WriteLine($"Replaced values in {fileName}. Saved as {outputFilePath}");

            }
            MessageBox.Show(numReplacments + " Registers replaced");
        }

        public void GenerateAllApps(string inputDirectoryPath, string outputDirectoryPath, string searchValue)
        {
            if (!Directory.Exists(outputDirectoryPath))
            {
                Directory.CreateDirectory(outputDirectoryPath);
            }


            byte searchInt = byte.Parse(searchValue);
            byte[] reaplaceApps = new byte[] { 4, 5, 6, 7, 8, 8, 9, 10, 11, 12, 13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,
            31,32,33,34,35,36,37,38,39,40, 41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,
            71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101,102,103,104,105,106,107,108,109,110,
            111,112,113,114,115,116,117,118,119,120,121,122,123,124,125,126,127,128,129,130,131,132,133,134,135,136,137,138,139,140,
            141,142,143,144,145,146,147,148,149,150,151,152,153,154,155,156,157,158,159,160,161,162,163,164,165,166,167,168,169,170,
            171,172,173,174,175,176,177,178,179,180,181,182,183,184,185,186,187,188,189,190,191,192,193,194,195,196,197,198,199,200,
            201,202,203,204,205,206,207,208,209,210,211,212,213,214,215,216,217,218,219,220,221,222,223,224,225,226,227,228,229,230,
            231,232,233,234,235,236,237,238,239,240,241,242,243,244,245,246,247,248,249,250,251,252,253};
            int numReplacments = 0;

            List<string> fileNames = new List<string>();
            fileNames.AddRange(Directory.GetFiles(inputDirectoryPath));

            foreach (byte newApp in reaplaceApps)
            { 
                //Per Group of files
                string folderName = "App " + newApp.ToString() + " " + inputDirectoryPath.Split('\\').Last();
                Directory.CreateDirectory(outputDirectoryPath + "\\" + folderName);

                foreach (string inputFilePath in fileNames)
                {
                    //Per File
                    string fileName = Path.GetFileName(inputFilePath);
                    string outputFilePath = Path.Combine(outputDirectoryPath + "\\" + folderName, fileName);
                    byte[] fileContent = File.ReadAllBytes(inputFilePath);
                    int skipByte = 0;
                    for (int i = 0; i < fileContent.Length - 4; i++)
                    {
                        if (skipByte > 0)
                        {
                            skipByte--;
                            continue;
                        }
                        byte[] word = { fileContent[i], fileContent[i + 1], fileContent[i + 2], fileContent[i + 3] };
                        byte currentValue = word[0];
                        if (currentValue == searchInt)
                        {
                            fileContent[i] = newApp;

                            skipByte = 3;
                            numReplacments++;
                        }
                    }

                    File.WriteAllBytes(outputFilePath, fileContent);
                    Console.WriteLine($"Replaced values in {fileName}. Saved as {outputFilePath}");

                }
            }
 
            MessageBox.Show(numReplacments + " Registers replaced");
        }
        public void ReplaceMultipleRegister(string inputDirectoryPath, string outputDirectoryPath, string findReplaceTemplate)
        {
            if (!Directory.Exists(outputDirectoryPath))
            {
                Directory.CreateDirectory(outputDirectoryPath);
            }

            CSVReader cSVReader = new CSVReader();
            cSVReader.ParseCSV(findReplaceTemplate, ",");

            List<PCCURegister> searchRegisters = new List<PCCURegister>();
            List<PCCURegister> replaceRegisters = new List<PCCURegister>();
            foreach (List<string> rows in cSVReader.GetArrayStorage())
            {
                searchRegisters.Add(new PCCURegister(rows[0]));
                replaceRegisters.Add(new PCCURegister(rows[1]));
            }

            List<UInt32> searchRegisterValues = new List<UInt32>();
            foreach (PCCURegister pCCURegister in searchRegisters)
            {
                searchRegisterValues.Add(pCCURegister.ToInt32());
            }

            List<UInt32> replaceRegisterValues = new List<UInt32>();
            foreach (PCCURegister pCCURegister in replaceRegisters)
            {
                replaceRegisterValues.Add(pCCURegister.ToInt32());
            }


            int numReplacments = 0;

            List<string> fileNames = new List<string>();
            fileNames.AddRange(Directory.GetFiles(inputDirectoryPath));
            // Create a new StreamWriter object for the log file.
            using (StreamWriter writer = new StreamWriter(outputDirectoryPath + "\\" + "_log.txt"))
            {
                // Write some text to the file.
                writer.WriteLine("This is a log of every change made to each file.\n");

                foreach (string inputFilePath in fileNames)
                {
                    
                    string fileName = Path.GetFileName(inputFilePath);
                    writer.WriteLine("File: " + fileName + "\n");
                    string outputFilePath = Path.Combine(outputDirectoryPath, fileName);
                    byte[] fileContent = File.ReadAllBytes(inputFilePath);
                    int skipByte = 0;
                    for (int i = 0; i < fileContent.Length - 4; i++)
                    {
                        if (skipByte > 0)
                        {
                            skipByte--;
                            continue;
                        }
                        byte[] word = { fileContent[i], fileContent[i + 1], fileContent[i + 2], fileContent[i + 3] };
                        UInt32 currentValue = BitConverter.ToUInt32(word, 0);

                        for (int n = 0; n < searchRegisterValues.Count; n++)
                        {
                            if (currentValue == searchRegisterValues[n])
                            {
                                writer.WriteLine("Found: " + searchRegisters[n].ToRegString() + " | Replaced: " + replaceRegisters[n].ToRegString() + "\n");
                                byte[] replaceByte = BitConverter.GetBytes(replaceRegisterValues[n]);
                                fileContent[i] = replaceByte[0];
                                fileContent[i + 1] = replaceByte[1];
                                fileContent[i + 2] = replaceByte[2];
                                fileContent[i + 3] = replaceByte[3];
                                skipByte = 3;
                                numReplacments++;
                                break;
                            }
                        }

                    }
                    writer.WriteLine("\n");
                    File.WriteAllBytes(outputFilePath, fileContent);
                    Console.WriteLine($"Replaced values in {fileName}. Saved as {outputFilePath}");
                   
                }
                MessageBox.Show(numReplacments + " Registers replaced");
            }
        }
    }
}
