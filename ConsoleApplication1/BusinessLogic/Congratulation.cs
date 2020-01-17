using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1.BusinessLogic
{
    class Congratulation
    {
        private string name;

        private string wish1;
        private string wish2;
        private string wish3;

        public string Name
        {
            get { return name;}
            set { name = value;}
        }
        public string Wish3
        {
            get
            {
                return wish3;
            }

            set
            {
                wish3 = value;
            }
        }
        public string Wish2
        {
            get
            {
                return wish2;
            }

            set
            {
                wish2 = value;
            }
        }
        public string Wish1
        {
            get
            {
                return wish1;
            }

            set
            {
                wish1 = value;
            }
        }

        public Congratulation() { }
            
        public Congratulation(string name, string wish1, string wish2, string wish3)
        {
            this.name = name;
            this.wish1 = wish1;
            this.wish2 = wish2;
            this.wish3 = wish3;
        }



    }
}
