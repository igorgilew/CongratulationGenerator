using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ConsoleApplication1.BusinessLogic
{
    class ExcelWorker
    {
        private Application app;
        private Workbook workBook;
        private Worksheet workSheet;
        private int countPeoples = 10;

        private string path = @"C:\Практика\Алёнке\ConsoleApplication1\ConsoleApplication1\Pozhelania_i_adresaty.xlsx";

        public ExcelWorker()
        {
            app = new Application();

        }

        /// <summary>
        /// функция для загрузки листа
        /// </summary>
        /// <param name="sheetNum">номер листа</param>
        private void loadSheet(int sheetNum)
        {
            workBook = app.Workbooks.Open(path, 0, false, 5, "", "", false, 
                XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            workSheet = (Worksheet)workBook.Sheets[sheetNum];
        }

        private string getCell(string cell)
        {
            Range range = workSheet.Range[cell];
            return range.Text.ToString();
        }

        public List<string> getNamesOfAdressees()
        {
            loadSheet(1);
            var adressees = new List<string>();
            for (int i = 1; i<= countPeoples; i++)
            {
                adressees.Add(getCell("A" + i));
            }
            return adressees;
        }

        public void test()
        {
            loadSheet(1);
            string res = getCell("A1");
            Console.WriteLine(res);


        }

        public void closeApp()
        {
            app.Quit(); 
        }
    }
}
