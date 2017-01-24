
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace ConsoleApp3
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application excelapp;
            Excel.Workbooks excelappworkbooks;
            Excel.Workbook excelappworkbook;
            Excel.Sheets excelsheets;
            Excel.Worksheet excelworksheet;
            Excel.Range excelcells;
            Excel.Window excelWindow;

            /*Блок переменных*/
            int count = 6768; //количество записей в файле для обработки
            string[] stringArray = new string[count];
            int i; //счётчик


            /*Указание файла*/
            Console.WriteLine("Введите имя файла, без раcширения");
            string adr = "ADDRHOUSE_del_ADDRESSFULL"; //Console.ReadLine();
            excelapp = new Excel.Application();
            excelapp.Visible = true;
            //excelapp.Workbooks.Open(@"D:\\Users\\amarchenko\\Desktop\\IBS\\" + adr + ".xlsm");
            excelappworkbook = excelapp.Workbooks.Open(@"D:\\Users\\amarchenko\\Desktop\\IBS\\" + adr + ".xlsm");
            excelsheets = excelappworkbook.Worksheets;


            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

            for (i = 0; i < count; i++)
            {
                excelcells = excelworksheet.get_Range("AO" + (i + 3));
                stringArray[i] = excelcells.Value;
            }


            //            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(2);
            for (i = 0; i < count; i++)
            {
                if (stringArray[i] != null)
                {

                    //                excelcells = excelworksheet.get_Range("AP" + (i + 1));
                    //                excelcells.Value2 = stringArray[i];
                    /*индекс*
                                        string pattern1 = @"\b(\d+)";
                                        Regex regex1 = new Regex(pattern1);
                                        Match match1 = regex1.Match(stringArray[i]);
                                        if (match1.Groups[1].Value.Length > 4)
                                        {
                                            excelcells = excelworksheet.get_Range("AP" + (i + 3));
                                            excelcells.Value2 = match1.Groups[1].Value;
                                        }
                    /*страна*
                                        string pattern2 = @"(Россия|Украина|РФ|Российская Федерация)";
                                        Regex regex2 = new Regex(pattern2, RegexOptions.IgnoreCase);
                                        Match match2 = regex2.Match(stringArray[i]);
                                        excelcells = excelworksheet.get_Range("AQ" + (i + 3));
                                        excelcells.Value2 = match2.Groups[1].Value;
                    /*область*
                                        string pattern3 = @"(\w+\W\w+\sРеспублика|\w+\sРеспублика|Республика\s\w+|\w+\W+обл|\w+\W+край)";//@"([д]\W+\d+)";
                                        Regex regex3 = new Regex(pattern3, RegexOptions.IgnoreCase);
                                        Match match3 = regex3.Match(stringArray[i]);
                                        excelcells = excelworksheet.get_Range("AR" + (i + 3));
                                        excelcells.Value2 = match3.Groups[1].Value;
                    /*район*
                                        string pattern4 = @"(\w+\W+р-н|\w+\W+район)";//@"([д]\W+\d+)";
                                        Regex regex4 = new Regex(pattern4, RegexOptions.IgnoreCase);
                                        Match match4 = regex4.Match(stringArray[i]);
                                        excelcells = excelworksheet.get_Range("AS" + (i + 3));
                                        excelcells.Value2 = match4.Groups[1].Value;
                                        /*город*/
                    string pattern5 = @"\b(г.\W\w+)";//город\s\w+|город\w+|г.\s\w+|г.\w+|село\s\w+|пос.\s\w+|Москва)";
                    Regex regex5 = new Regex(pattern5, RegexOptions.IgnoreCase);
                    Match match5 = regex5.Match(stringArray[i]);
                    excelcells = excelworksheet.get_Range("AT" + (i + 3));
                    excelcells.Value2 = match5.Groups[1].Value;
                    /*улица*

string pattern6 = @"(\w+\sвал|\w+\sпроезд|ул.\w+|ул.\s\w+|\w+\sплощадь|\w+\sш.|\w+\sшоссе|\w+\sнаб.|\w+\sнабережная|бульвар\s\w+|\w+\sпер.|\w+\sпереулок|проспект\s\w+)";
                    Regex regex6 = new Regex(pattern6, RegexOptions.IgnoreCase);
                    Match match6 = regex6.Match(stringArray[i]);
                    excelcells = excelworksheet.get_Range("AU" + (i + 3));
                    excelcells.Value2 = match6.Groups[1].Value;
                    /*дом*
                    string pattern7 = @"(д.\d+\w|д.\s\d+\w|д.\s\d+|д.\d+|д.\d+\s\d+|вл.\s+\d+|вл.\d+|влад.\s+\d+|дом\s\d+\w|дом\s\d+)";
                    Regex regex7 = new Regex(pattern7, RegexOptions.IgnoreCase);
                    Match match7 = regex7.Match(stringArray[i]);
                    excelcells = excelworksheet.get_Range("AV" + (i + 3));
                    excelcells.Value2 = match7.Groups[1].Value;
                    /*корпус*
                    string pattern8 = @"(к\s\d+|к.\d+|к.\s\d+|корпус\s+\d+|корп.\s+\d+|помещение\s+\d+)";//|строение\s+\d+|стр.\s+\d+|стр.\d+)";
                    Regex regex8 = new Regex(pattern8, RegexOptions.IgnoreCase);
                    Match match8 = regex8.Match(stringArray[i]);
                    excelcells = excelworksheet.get_Range("AW" + (i + 3));
                    excelcells.Value2 = match8.Groups[1].Value;
                    /*строение*
                    string pattern9 = @"(строение\s+\d+|стр.\s+\d+|стр.\d+)";
                    Regex regex9 = new Regex(pattern9, RegexOptions.IgnoreCase);
                    Match match9 = regex9.Match(stringArray[i]);
                    excelcells = excelworksheet.get_Range("AX" + (i + 3));
                    excelcells.Value2 = match9.Groups[1].Value;
                    /*аппартаменты*
                    string pattern10 = @"(комн.\s\d+|комн.\d+|комн\s\d+|комн\d+|кв.\s\d+|кв.\d+|кв\s\d+|кв\d+|ком.\s\d+|ком.\d+|ком\s\d+|ком\d+|оф.\s\d+|оф.\d+|оф\s\d+|оф\d+)";
                    Regex regex10 = new Regex(pattern10, RegexOptions.IgnoreCase);
                    Match match10 = regex10.Match(stringArray[i]);
                    excelcells = excelworksheet.get_Range("AY" + (i + 3));
                    excelcells.Value2 = match10.Groups[1].Value;

                    //            excelapp.Quit();  /**/
                }
            }
        }
    }
}