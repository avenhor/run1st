using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.VisualBasic.FileIO;
using System.Threading.Tasks;
using System.Globalization;

namespace GUID_converter
{
    class Program
    {
        static void Main(string[] args)
        {
            //printFile();
            writeData();
            Console.ReadKey();
        }

        public static byte[] ConvertHexStringToByteArray(string hexString)
        {
            if (hexString.Length % 2 != 0)
            {
                throw new ArgumentException(String.Format(CultureInfo.InvariantCulture, "The binary key cannot have an odd number of digits: {0}", hexString));
            }

            byte[] HexAsBytes = new byte[hexString.Length / 2];
            for (int index = 0; index < HexAsBytes.Length; index++)
            {
                string byteValue = hexString.Substring(index * 2, 2);
                HexAsBytes[index] = byte.Parse(byteValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            }
            return HexAsBytes;
        }

        public static void printFile()
        {
            using (TextFieldParser parser = new TextFieldParser(@"C:\Users\jmcarthur\Documents\Projects\PD AD Export\pd_employee.csv"))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                int counter = 0;
                int ptr = 4;
                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();
                    foreach (string field in fields)
                    {
                        counter += 1;
                        Console.WriteLine(String.Format("counter {0}: {1}", counter, field));
                    }
                    if (counter > 20) { break; }
                }
            }
        }

        public static void writeData()
        {
            string fileName = @"C:\Users\jmcarthur\Documents\Projects\PD AD Export\EMPLOYEE.csv";
            using (TextFieldParser parser = new TextFieldParser(@"C:\Users\jmcarthur\Documents\Projects\PD AD Export\pd_employee.csv"))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                int counter = 0;
                int ptr = 4;
                StreamWriter sw = new StreamWriter(fileName);
                while (!parser.EndOfData)
                {
                    //Process row
                    string[] fields = parser.ReadFields();
                    try
                    {
                        foreach (string field in fields)
                        {
                            counter += 1;
                            if (counter < 7)
                            {
                                if (counter % 6 == 0)
                                {
                                    sw.WriteLine(field);
                                }
                                else
                                {
                                    sw.Write(field + ",");
                                }
                            }
                            else
                            {
                                if ((counter - 6) == ptr)
                                {
                                    Console.WriteLine("Trying to convert {0} to byte[]", field);
                                    byte[] hexString = ConvertHexStringToByteArray(field);
                                    Guid guid = new Guid(hexString);
                                    sw.Write(guid.ToString() + ",");
                                    Console.WriteLine(String.Format("\tGUID: {0}", guid.ToString()));
                                    ptr = counter;
                                }
                                else
                                {
                                    if (counter % 6 == 0)
                                    {
                                        sw.WriteLine(field);
                                    }
                                    else
                                    {
                                        sw.Write(field + ",");
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(String.Format("Caught exception writing data to file: {0}", e));
                    }
                }
                sw.Close();
                sw.Dispose();
            }
        }
    }
}
