using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using ConsoleApplication1.BusinessLogic;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Congratulation> congratulationList = new List<Congratulation>();
            //Congratulation c1 = new BusinessLogic.Congratulation("Игорь", "счастья", "здоровья", "успехов");
            //Congratulation c2 = new BusinessLogic.Congratulation("Игорь2", "счастья", "здоровья", "успехов");

            //congratulationList.Add(c1);
            //congratulationList.Add(c2);

            WordWorker ww = new WordWorker();
            
            ExcelWorker ew = new ExcelWorker();
            var adressees = new List<string>();
            adressees = ew.getNamesOfAdressees();
            foreach(var name in adressees)
            {
                Congratulation c1 = new BusinessLogic.Congratulation(name, "счастья", "здоровья", "успехов");
                congratulationList.Add(c1);
            }
            ew.closeApp();

            foreach (var c in congratulationList)
            {
                ww.createCongratulation(c);
                
            }

            ww.showApp();
            ww.closeApp();

            Console.ReadLine();

        }
    }
}
