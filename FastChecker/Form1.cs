using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace FastChecker
{
    public partial class Form1 : Form
    {

        private static Excel.Application excel;
        private static Excel.Workbooks excelWorkbooks;
        private static Excel.Workbook excelWorkbook;
        private static Excel.Sheets excelSheets;
        private static Excel.Worksheet excelWorksheet;
        private static Excel.Range excelRange;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            excel = new Excel.Application(); // запуск приложения 

            excel.Visible = true; // окно видимо

            excelWorkbook = excel.Workbooks.Open("D:\\Users\\DSIYANCHEV\\Desktop\\ADDRESSFULL_разбиение.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); // открыть книгу

            excelSheets = excelWorkbook.Worksheets; //выгрузить листы 

            excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(1); //получение ссылки на лист


            for (int i = 2258; i < 4515; i ++)
            {
                try
                {
                    string index = excelWorksheet.get_Range("P" + i.ToString(), Type.Missing).Value2.ToString();
                    string country = excelWorksheet.get_Range("R" + i.ToString(), Type.Missing).Value2.ToString();
                    string region = excelWorksheet.get_Range("T" + i.ToString(), Type.Missing).Value2.ToString();
                    string city = excelWorksheet.get_Range("AD" + i.ToString(), Type.Missing).Value2.ToString();
                    string district = excelWorksheet.get_Range("X" + i.ToString(), Type.Missing).Value2.ToString();
                    string street = excelWorksheet.get_Range("AH" + i.ToString(), Type.Missing).Value2.ToString();
                    string house = excelWorksheet.get_Range("AJ" + i.ToString(), Type.Missing).Value2.ToString();
                    string building = excelWorksheet.get_Range("AK" + i.ToString(), Type.Missing).Value2.ToString();
                    string housing = excelWorksheet.get_Range("AL" + i.ToString(), Type.Missing).Value2.ToString();
                    string apartment = excelWorksheet.get_Range("AN" + i.ToString(), Type.Missing).Value2.ToString();
                    string full = excelWorksheet.get_Range("AO" + i.ToString(), Type.Missing).Value2.ToString();
                    dataGridView1.Rows.Add(i.ToString(), index, country, region, district, city, street, house, building, housing, apartment, full);
                }
                catch
                {
                    continue;
                }
            }

            excel.Quit();
        }
    }
}
