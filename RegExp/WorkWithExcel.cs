﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace RegExp
{
    class WorkWithExcel
    {
        private static Excel.Application excel;
        private static Excel.Workbooks excelWorkbooks;
        private static Excel.Workbook excelWorkbook;
        private static Excel.Sheets excelSheets;
        private static Excel.Worksheet excelWorksheet;
        private static Excel.Range excelRange;

        public static List<string> readAllLines(string path, int page, string column)
        {
            List<string> output = new List<string>();

            excel = new Excel.Application(); // запуск приложения 

            excel.Visible = true; // окно видимо

            excelWorkbook = excel.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); // открыть книгу

            excelSheets = excelWorkbook.Worksheets; //выгрузить листы 

            excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(page); //получение ссылки на лист

            int i = 1; // итератор (номер строки)
            string cell = column + i.ToString(); // ячейка, откуда берется информация 
            while ((excelRange = excelWorksheet.get_Range(cell, Type.Missing)).ToString().Length != 0) // пока ячейка не пустая
            {
                try
                {
                    output.Add(excelRange.Value2.ToString());
                    i++;
                    cell = column + i.ToString();
                }
                catch
                {
                    break;
                }
                
            }

            excel.Quit();
            return output;
        }

        public static void putColumns(string path, int page, List<string> columns, List<List<string>> data)
        {
            excel = new Excel.Application(); // запуск приложения 

            excel.Visible = true; // окно видимо

            excelWorkbook = excel.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing); // открыть книгу

            excelSheets = excelWorkbook.Worksheets; //выгрузить листы 

            excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(page); //получение ссылки на лист

            for (int i = 0; i < columns.Count; i ++)
            {
                for (int j = 0; j < data[i].Count; j ++)
                {
                    string cell = columns[i] + (j + 1).ToString();
                    excelRange = excelWorksheet.get_Range(cell, Type.Missing);
                    excelRange.Value2 = data[i][j];
                }
            }

            excelWorkbook.Save();
            excel.Quit();
        }
    }
}
