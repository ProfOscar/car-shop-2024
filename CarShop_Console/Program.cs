using CarShop_Library;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CarShop_Console
{
    internal class Program
    {
        static List<Veicolo> ParcoMezzi = new List<Veicolo>();

        static void Main(string[] args)
        {
            CreaDatiDiProva();
            char scelta = ' ';
            while (scelta != 'q' && scelta != 'Q')
            {
                scelta = ScriviMenu();
                switch (scelta)
                {
                    case 'c':
                    case 'C':
                        CaricaDati();
                        break;
                    case 's':
                    case 'S':
                        SalvaDati();
                        break;
                    case '1':
                        Elenco<Veicolo>();
                        break;
                    case '2':
                        Elenco<Auto>();
                        break;
                    case '3':
                        Elenco<Moto>();
                        break;
                    default:
                        break;
                }
            }
        }

        private static void CaricaDati()
        {
            try
            {
                ParcoMezzi = JsonTools.CaricaDati();
            }
            catch (Exception exc)
            {
                Console.WriteLine("\nEccezione in caricamento da file json: " + exc.Message);
                Console.ReadKey(true);
                Console.Clear();
            }
        }

        private static void SalvaDati()
        {
            if (JsonTools.SalvaDati(ParcoMezzi))
                Console.WriteLine("\n*** SCRITTURA DATI OK ***");
            else
                Console.WriteLine("\n*** PROBLEMI SU SCRITTURA DATI ***");
            // Thread.Sleep(2000);
            Console.ReadKey(true);
            Console.Clear();
        }

        private static char ScriviMenu()
        {
            Console.Clear();
            Console.WriteLine("*** GESTIONE RIVENDITA VEICOLI ***");
            Console.WriteLine("".PadLeft(34, '-'));
            Console.WriteLine("C - CARICA Dati");
            Console.WriteLine("S - SALVA Dati");
            Console.WriteLine("".PadLeft(34, '-'));
            Console.WriteLine("1 - Visualizza TUTTI");
            Console.WriteLine("2 - Visualizza AUTO");
            Console.WriteLine("3 - Visualizza MOTO");
            Console.WriteLine("".PadLeft(34, '-'));
            Console.WriteLine("Q - ESCI");
            return Console.ReadKey(true).KeyChar;
        }

        private static void Elenco<T>()
        {
            Console.Clear();
            string typeName = typeof(T).Name.ToUpper();
            Console.WriteLine($"*** ELENCO {typeName} ***");
            int conta = 0;
            List<Veicolo> filteredItems = ParcoMezzi.FindAll(element => element is T);
            foreach (var item in filteredItems)
            {
                conta++;
                Console.WriteLine($"\n{conta} - {item}");
            }
            Console.WriteLine($"\n\nTOT: {conta} {typeName}");
            Console.ReadKey(true);
        }

        private static void CreaDatiDiProva()
        {
            try
            {
                ParcoMezzi = JsonTools.CaricaDati();
            }
            catch (Exception exc)
            {
                Veicolo v = new Auto("BMW", "Serie 3", "FE518TW", new DateTime(2015, 10, 24), 19000, "",
                    TipoAlimentazione.Benzina, 5, true);
                ParcoMezzi.Add(v);
                v = new Auto("Mercedes", "CLA", "GA331LD", new DateTime(2022, 2, 15), 38000, "",
                    TipoAlimentazione.Diesel, 5, false);
                ParcoMezzi.Add(v);

                v = new Moto("Ducati", "Diavolo Rosso", "EH654TY", new DateTime(2022, 7, 10), 8500, "",
                    TipoMoto.Enduro, 4, false);
                ParcoMezzi.Add(v);

                v = new Furgone("FIAT", "Ducato", "E451UF", new DateTime(2007, 3, 11), 4200, "", 12000);
                ParcoMezzi.Add(v);

                Console.WriteLine("Eccezione in caricamento da file json: " + exc.Message);
                Console.WriteLine("Dati di prova creati staticamente");
                Console.ReadKey(true);
                Console.Clear();
            }
        }
    }
}
