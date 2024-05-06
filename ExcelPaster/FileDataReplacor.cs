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
