using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using _Execl = Microsoft.Office.Interop.Excel;

namespace ConsoleApp7
{
    class Program
    {
        static void Main(string[] args)
        {
            _Application excel = new _Execl.Application();
            Workbook wb;
            Worksheet ws;
            int i = 1, j = 1;
            wb = excel.Workbooks.Open(@"C:\Users\kuche\Desktop\Учеба 2 курс\Разработка кода информационных систем\2 курс\c#\ConsoleApp7\Лист Microsoft Excel1.xlsx");
            ws = wb.Worksheets[1];

            if (ws.Cells[i, j].Value2 != null)
                MessageBox.Show(ws.Cells[i, j].Value);
            else
                MessageBox.Show("");
            wb.Close();

            wb = excel.Workbooks.Open(@"C:\Users\kuche\Desktop\Учеба 2 курс\Разработка кода информационных систем\2 курс\c#\ConsoleApp7\Лист Microsoft Excel1.xlsx");
            ws = wb.Worksheets[1];

            Console.WriteLine("Введите новый текст");
            string a = Console.ReadLine();
            ws.Cells[i, j].Value2 = a;
            wb.SaveAs(@"C:\Users\kuche\Desktop\Учеба 2 курс\Разработка кода информационных систем\2 курс\c#\ConsoleApp7\Лист Microsoft Excel1.xlsx");
            wb.Close();
        }
    }
}
