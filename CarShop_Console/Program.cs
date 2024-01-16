using CarShop_Library;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CarShop_Console
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // string myName = "Oscar";
            // Console.WriteLine(ClassTest.HelloWorld(myName));

            Veicolo v = new Auto("BMW", "Serie 3", "FE518TW", new DateTime(2015, 10, 24), 19000, 5, true);
            Console.WriteLine(v);

            Console.WriteLine();

            v = new Moto("Ducati", "Diavolo Rosso", "EH654TY", new DateTime(2022, 7, 10), 8500, 4, false);
            Console.WriteLine(v);

            Console.WriteLine();

            v = new Furgone("FIAT", "Ducato", "E451UF", new DateTime(2007, 3, 11), 4200, 12000);
            Console.WriteLine(v);

            Console.ReadLine();
        }
    }
}
