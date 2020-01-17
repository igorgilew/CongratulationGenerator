using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1.BusinessLogic
{
    class WordWorker
    {
        private string path = @"C:\Практика\Алёнке\ConsoleApplication1\ConsoleApplication1\Shablon.docx";
        private Application app;
        private Document doc = null;
        List<string> bookmarks = new List<string>() { "name", "wish1", "wish2", "wish3" };

        public WordWorker()
        {
            app = new Application();
            openDocument();
        }

        private void openDocument()
        {
            try
            {
                doc = app.Documents.Add(path);
            }
            catch (Exception e)
            {
                Console.WriteLine("Не удалось открыть шаблон");
                Console.WriteLine(e.Message);
                Console.ReadKey();
                app.Quit();
                return;
            }
        }

        public void createCongratulation(Congratulation cngrtln)
        {           

            var bm = app.ActiveDocument.Bookmarks[bookmarks[0]].Range.Text=cngrtln.Name;
            bm = app.ActiveDocument.Bookmarks[bookmarks[1]].Range.Text = cngrtln.Wish1;
            bm = app.ActiveDocument.Bookmarks[bookmarks[2]].Range.Text = cngrtln.Wish2;
            bm = app.ActiveDocument.Bookmarks[bookmarks[3]].Range.Text = cngrtln.Wish3;

            app.Selection.EndKey(WdUnits.wdStory);
            app.Selection.InsertNewPage();
            app.Selection.InsertFile(path, "", true, false, false);

        }



        public void showApp()
        {
            Console.WriteLine("Генерация завершена!");
            
            app.Visible = true;
            Console.ReadKey();
        }


        public void closeApp()
        {
            app.Quit();
        }


    }
}
