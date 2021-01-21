
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
// JSON
using System.Text.Json;
using Newtonsoft.Json;
// Unicode --> Chinese Characters
using System.Text.Encodings.Web;
using System.Text.Unicode;





namespace lotteryPickNPOI2
{
    public partial class Form1 : Form
    {

        List<Person> personNotSelected = new List<Person>();
        List<Person> personSelected = new List<Person>();

        List<int> indexNotSelected = new List<int>();


        List<int> indexSelected = new List<int>();

        Random random = new Random();

        private readonly string _path = @"C:\Users\hankm\OneDrive\Documents\Aurocore-DESKTOP-U594CEO-DESKTOP-U594CEO\lotteryPickNPOI3-Final\lotteryResult.json";
        // JsonResultHTTP
        private readonly string _path2 = @"C:\Users\hankm\OneDrive\Documents\Aurocore-DESKTOP-U594CEO-DESKTOP-U594CEO\jsonResultHTTP\lotteryResult.json";
        // lotteryPickWebsite2
        private readonly string _path3 = @"C:\Users\hankm\OneDrive\Documents\Aurocore-DESKTOP-U594CEO-DESKTOP-U594CEO\lotteryPickWebsite2\lotteryResult.json";

        // Original lotteryList JSON file path
        private readonly string _path4 = @"C:\Users\hankm\OneDrive\Documents\Aurocore-DESKTOP-U594CEO-DESKTOP-U594CEO\originalLotteryList.json";

        public Form1()
        {
            InitializeComponent();
            
        }


        private void btnStartLottery_Click(object sender, EventArgs e)
        {
            String filePath = @"C:\Users\hankm\OneDrive\Documents\Aurocore-DESKTOP-U594CEO-DESKTOP-U594CEO\lotteryList.xlsx";
            //try
            //{
                IWorkbook workbook = null;
                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                if (filePath.IndexOf(".xlsx") > 0)
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (filePath.IndexOf(".xls") > 0)
                {
                    workbook = new HSSFWorkbook(fs);
                }

                ISheet sheet = workbook.GetSheetAt(0);
                
                if (sheet != null)
                {
                    int rowCount = sheet.LastRowNum;
                    // Console.WriteLine(rowCount);

                    for (int i = 1; i <= rowCount; i++)
                    {
                        /*
                            if (i == 5181)
                            {
                                Console.WriteLine("AHA GOTEM");
                            }
                        */
                            IRow curRow = sheet.GetRow(i);

                        // Time Stamp
                        string cellValue0 = "";
                        if (curRow.GetCell(0).CellType.ToString() == "String")
                        {
                            cellValue0 = curRow.GetCell(0).StringCellValue.Trim();
                        }
                        else
                        {
                            cellValue0 = curRow.GetCell(0).NumericCellValue.ToString().Trim();
                        }

                        //發票號碼
                        string cellValue1 = "";
                        if (curRow.GetCell(1).CellType.ToString() == "String")
                        {
                            cellValue1 = curRow.GetCell(1).StringCellValue.Trim();
                        }
                        else
                        {
                            cellValue1 = curRow.GetCell(1).NumericCellValue.ToString().Trim();
                        }

                        // 消費金額
                        string cellValue2 = "";
                        if (curRow.GetCell(2).CellType.ToString() == "String")
                        {
                            cellValue2 = curRow.GetCell(2).StringCellValue.Trim();
                        }
                        else
                        {
                            cellValue2 = curRow.GetCell(2).NumericCellValue.ToString().Trim();
                        }

                        // 縣市
                        string cellValue3 = "";
                        if (curRow.GetCell(3).CellType.ToString() == "String")
                        {
                            cellValue3 = curRow.GetCell(3).StringCellValue.Trim();
                        }
                        else
                        {
                            cellValue3 = curRow.GetCell(3).NumericCellValue.ToString().Trim();
                        }

                        // 區
                        string cellValue4 = "";
                        if (curRow.GetCell(4).CellType.ToString() == "String")
                        {
                            cellValue4 = curRow.GetCell(4).StringCellValue.Trim();
                        }
                        else
                        {
                            cellValue4 = curRow.GetCell(4).NumericCellValue.ToString().Trim();
                        }

                        //門市
                        string cellValue5 = "";
                        if (curRow.GetCell(5).CellType.ToString() == "String")
                        {
                            cellValue5 = curRow.GetCell(5).StringCellValue.Trim();
                        }
                        else
                        {
                            cellValue5 = curRow.GetCell(5).NumericCellValue.ToString().Trim();
                        }

                        // 姓名
                        string cellValue6 = "";
                        if (curRow.GetCell(6).CellType.ToString() == "String")
                        {
                            cellValue6 = curRow.GetCell(6).StringCellValue.Trim();
                        }
                        else
                        {
                            cellValue6 = curRow.GetCell(6).NumericCellValue.ToString().Trim();
                        }
                        //電話
                        string cellValue7 = "";
                        if (curRow.GetCell(7).CellType.ToString() == "String")
                        {
                            cellValue7 = curRow.GetCell(7).StringCellValue.Trim();
                        }
                        else
                        {
                            cellValue7 = curRow.GetCell(7).NumericCellValue.ToString().Trim();
                        }

                        //電子郵件
                        string cellValue8 = "";
                        if (curRow.GetCell(8).CellType.ToString() == "String")
                        {
                            cellValue8 = curRow.GetCell(8).StringCellValue.Trim();
                        }
                        else
                        {
                            cellValue8 = curRow.GetCell(8).NumericCellValue.ToString().Trim();
                        }

                        Console.WriteLine(cellValue0 + "\t\t" + cellValue1 + "\t\t" + cellValue2 + "\t\t" + cellValue3 + "\t\t" + cellValue4 + "\t\t" + cellValue5 + "\t\t" + cellValue6 + "\t\t" + cellValue7 + "\t\t" + cellValue8 + "\n\n");

                        personNotSelected.Add(new Person(cellValue0, cellValue1, cellValue2, cellValue3, cellValue4, cellValue5, cellValue6, cellValue7, cellValue8));
                    }

                // Get a json file with original lottery list (NOT TOUCHED AT ALL)
                JsonSerializerOptions optionsOther = new JsonSerializerOptions
                {
                    Encoder = JavaScriptEncoder.Create(UnicodeRanges.CjkUnifiedIdeographs, UnicodeRanges.CjkUnifiedIdeographsExtensionA, UnicodeRanges.CjkCompatibilityIdeographs, UnicodeRanges.BasicLatin),
                    WriteIndented = true
                };
                var lotteryList = System.Text.Json.JsonSerializer.Serialize(personNotSelected, optionsOther);



                using (var writer = new StreamWriter(_path4))
                {
                    // string jsonstr = JsonConvert.SerializeObject(lotteryResultJson, new JsonSerializerSettings() { StringEscapeHandling = StringEscapeHandling.Default });
                    writer.Write(lotteryList);
                }
                // }
            }
            /*
            catch(Exception exception)
            {
                Console.WriteLine(exception.Message);
            }
            */

            // EITHER WAY, in total they have to add up to however many people need to be selected for prizes
            // 頭獎: 1
            // 二獎: 3
            // 三獎: 3
            // 四獎: 10
            // 購物袋: 3000
            
            
            for (int i = 1; i < 5182; i++)
            {
                indexNotSelected.Add(i);
            }

            for (int i = 0; i < 3017; i++)
            {
                int rnd = GenerateRandomNum(indexNotSelected.Count-1); // Generate a random number from 1 to the size of indexNotSelected
                int num = indexNotSelected[rnd]; // the index chosen from indexNotSelected
                indexSelected.Add(num); // Add the chose index to indexSelected
                indexNotSelected.RemoveAt(rnd); // Remove that chosen index in indexNotSelected

                // 
                for (int k = 0; k < indexNotSelected.Count; k++)
                {
                    // Delete the entries by the same person once he/she is selected for a prize
                    if (personNotSelected[num].phoneNumber == personNotSelected[k].phoneNumber && personNotSelected[num].name == personNotSelected[k].name && num != k)
                    {
                        indexNotSelected.Remove(k);
                    }
                }
            }

            for (int i = 0; i < indexSelected.Count; i++)
            {
                personSelected.Add(personNotSelected[indexSelected[i]]);
            }

            Console.WriteLine("\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\");
            Console.WriteLine("\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\");
            Console.WriteLine("\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\");

            for (int i = 0; i < personSelected.Count; i++)
            {
                printPersonAttributes(personSelected[i]);
            }


            // Write the results into a JSON file
            var lotteryResultJson = System.Text.Json.JsonSerializer.Serialize(personSelected);



            JsonSerializerOptions options = new JsonSerializerOptions
            {
                Encoder = JavaScriptEncoder.Create(UnicodeRanges.CjkUnifiedIdeographs, UnicodeRanges.CjkUnifiedIdeographsExtensionA, UnicodeRanges.CjkCompatibilityIdeographs, UnicodeRanges.BasicLatin),
                WriteIndented = true
            };
            var jsonString = System.Text.Json.JsonSerializer.Serialize(personSelected, options);



            using (var writer = new StreamWriter(_path)) 
            {
                // string jsonstr = JsonConvert.SerializeObject(lotteryResultJson, new JsonSerializerSettings() { StringEscapeHandling = StringEscapeHandling.Default });
                writer.Write(jsonString);
            }

            using (var writer = new StreamWriter(_path2))
            {
                writer.Write(jsonString);
            }

            using (var writer = new StreamWriter(_path3))
            {
                writer.Write(jsonString);
            }


        }

        public int GenerateRandomNum(int maxIndex)
        {
            // Create an GLOBAL array/ or multiple 
            Random random = new Random();

            int num = random.Next(1, maxIndex);  // Generate number between 1 and 5181

            return num;
        }

        public void printPersonAttributes(Person person)
        {
            Console.Write(person.timeStamp + "\t\t");
            Console.Write(person.receiptNum + "\t\t");
            Console.Write(person.purchaseAmount + "\t\t");
            Console.Write(person.city + "\t\t");
            Console.Write(person.area + "\t\t");
            Console.Write(person.location + "\t\t");
            Console.Write(person.name + "\t\t");
            Console.Write(person.phoneNumber + "\t\t");
            Console.WriteLine(person.email);
        }

        public static string Utf16ToUtf8(string utf16String)
        {
            /**************************************************************
             * Every .NET string will store text with the UTF16 encoding, *
             * known as Encoding.Unicode. Other encodings may exist as    *
             * Byte-Array or incorrectly stored with the UTF16 encoding.  *
             *                                                            *
             * UTF8 = 1 bytes per char                                    *
             *    ["100" for the ansi 'd']                                *
             *    ["206" and "186" for the russian 'κ']                   *
             *                                                            *
             * UTF16 = 2 bytes per char                                   *
             *    ["100, 0" for the ansi 'd']                             *
             *    ["186, 3" for the russian 'κ']                          *
             *                                                            *
             * UTF8 inside UTF16                                          *
             *    ["100, 0" for the ansi 'd']                             *
             *    ["206, 0" and "186, 0" for the russian 'κ']             *
             *                                                            *
             * We can use the convert encoding function to convert an     *
             * UTF16 Byte-Array to an UTF8 Byte-Array. When we use UTF8   *
             * encoding to string method now, we will get a UTF16 string. *
             *                                                            *
             * So we imitate UTF16 by filling the second byte of a char   *
             * with a 0 byte (binary 0) while creating the string.        *
             **************************************************************/

            // Storage for the UTF8 string
            string utf8String = String.Empty;

            // Get UTF16 bytes and convert UTF16 bytes to UTF8 bytes
            byte[] utf16Bytes = Encoding.Unicode.GetBytes(utf16String);
            byte[] utf8Bytes = Encoding.Convert(Encoding.Unicode, Encoding.UTF8, utf16Bytes);

            // Fill UTF8 bytes inside UTF8 string
            for (int i = 0; i < utf8Bytes.Length; i++)
            {
                // Because char always saves 2 bytes, fill char with 0
                byte[] utf8Container = new byte[2] { utf8Bytes[i], 0 };
                utf8String += BitConverter.ToChar(utf8Container, 0);
            }

            // Return UTF8
            return utf8String;
        }

        public void GeneratePick()
        {
            // EITHER WAY, in total they have to add up to however many people need to be selected for prizes
            // 頭獎: 1
            // 二獎: 3
            // 三獎: 3
            // 四獎: 10
            // 購物袋: 3000
            List<int> array = new List<int>();

            Random random = new Random();

            for (int i = 0; i < 3017; i++)
            {
                
            }

            // Generate random numbers from 1 to 5181 and store them into the array
            // CHECK FOR DUPLICATES, CAN'T HAVE ANY DUPLICATES (Meaning I gotta check customer's other attributes)
            // Attributes can PROBABLY JUST BE (Phone number OR email) (They're most likely unique anyways)


            // Return an array/list
        }
    }
}
