using CarShop_Library;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
                    case 'h':
                    case 'H':
                        EsportaHtml();
                        break;
                    case 'w':
                    case 'W':
                        EsportaDocx();
                        break;
                    default:
                        break;
                }
            }
        }

        private static void EsportaHtml()
        {
            int num;
            do
            {
                Console.Write("\nInserisci il numero del veicolo: ");
            } while (!int.TryParse(Console.ReadLine(), out num) || num <= 0 || num > ParcoMezzi.Count);
            Veicolo v = ParcoMezzi[num - 1];
            // copiamo template html su index.html
            File.Copy("./html/template.html", $"./html/{v.Targa}.html", true);
            // innestiamo in {v.Targa}.html i dati del veicolo richiesto
            string content = File.ReadAllText($"./html/{v.Targa}.html");
            content = content.Replace("{{marca}}", v.Marca).Replace("{{modello}}", v.Modello);
            // salviamo e apriamo nel browser di default il file {v.Targa}.html
            File.WriteAllText($"./html/{v.Targa}.html", content);
            string pgmDir = AppDomain.CurrentDomain.BaseDirectory;
            Process.Start($"{pgmDir}/html/{v.Targa}.html");
        }

        private static void EsportaDocx()
        {
            int num;
            do
            {
                Console.Write("\nInserisci il numero del veicolo: ");
            } while (!int.TryParse(Console.ReadLine(), out num) || num <= 0 || num > ParcoMezzi.Count);
            Veicolo v = ParcoMezzi[num - 1];
            // creiamo il volantino docx tramite OpenXML
            string pgmDir = AppDomain.CurrentDomain.BaseDirectory;
            string filePath = $"{pgmDir}/{v.Targa}.docx";
            TestVolantinoDocx(filePath);
            Process.Start(filePath);
        }

        private static void CaricaDati()
        {
            try
            {
                ParcoMezzi = JsonTools.CaricaDati();
                Console.WriteLine("\n*** CARICAMENTO DATI OK ***");
            }
            catch (Exception exc)
            {
                Console.WriteLine("\nEccezione in caricamento da file json: " + exc.Message);
            }
            Console.ReadKey(true);
            Console.Clear();
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
            Console.WriteLine("H - Crea Volantino HTML");
            Console.WriteLine("W - Crea Volantino DOCX");
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
                Console.WriteLine("\n*** CARICAMENTO DATI OK ***");
                Console.ReadKey(true);
                Console.Clear();
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

        public static void TestVolantinoDocx(string filePath)
        {
            using (WordprocessingDocument wordDocument = OpenXmlTools.CreaDocumento(filePath))
            {
                // prendo un riferimento al body del documento
                Body docBody = wordDocument.MainDocumentPart.Document.Body;

                // Aggiungo la stringa per i paragrafi di test
                string lorem = @"Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed aliquet mauris in magna finibus, ut porttitor felis condimentum. Cras sed hendrerit ex. Sed porta dictum purus eu dictum. Donec hendrerit aliquet mollis. Maecenas volutpat lacus eu lorem porta, quis imperdiet nibh pharetra. Sed ac eros diam. Sed ex libero, commodo in iaculis nec, scelerisque in erat. Proin ultricies hendrerit volutpat. Vivamus porttitor, nibh in maximus gravida, enim arcu porta leo, ac porttitor enim elit vel sapien.";

                // definisco stile
                Style myStyle = OpenXmlTools.CreaStile(wordDocument, "Codice 1", "left", "CCCCCC", "Courier New", 10, 50, 100);

                // utilizzo stile su paragrafo
                docBody.Append(OpenXmlTools.CreaParagrafoConStile(lorem, myStyle.StyleId));

                // test elenchi
                string[] contenutoElenchi = { "BMW Serie 3", "Jeep Compass", "Mercedes CLA", "Fiat Panda" };
                List<Paragraph> elenco = OpenXmlTools.CreaElenco(contenutoElenchi);
                foreach (var item in elenco) docBody.Append(item);

                // test tabella
                string[,] contenutoTabella = {
                    { "MARCA", "MODELLO", "TARGA", "PREZZO" },
                    { "BMW", "iX2", "GG528YT", "€ 57.800" },
                    { "Jeep", "Compass", "FR508HD", "€ 35750" }
                };
                Table table = OpenXmlTools.CreaTabella(contenutoTabella, "center", "right",
                    "red", "green",
                    380);
                docBody.Append(table);

                // 3 paragrafi semplici con diversa giustificazione
                docBody.Append(OpenXmlTools.CreaParagrafo(lorem));
                docBody.Append(OpenXmlTools.CreaParagrafo(lorem, "center"));
                docBody.Append(OpenXmlTools.CreaParagrafo(lorem, "right"));
                docBody.Append(OpenXmlTools.CreaParagrafo(lorem, "distribute"));

                // 1 paragrafo formattato in modo omogeneo
                docBody.Append(OpenXmlTools.CreaParagrafo(lorem, "left", false, true, false, "77FF33", "Tahoma", 15));

                // un paragrafo con il contenuto formattato nei diversi run
                Paragraph p = OpenXmlTools.CreaParagrafo("", "center");
                Run r = OpenXmlTools.CreaRun("Testo normale"); p.Append(r);
                r = OpenXmlTools.CreaRun("Testo grassetto", true); p.Append(r);
                r = OpenXmlTools.CreaRun("Testo corsivo", false, true); p.Append(r);
                r = OpenXmlTools.CreaRun("Testo sottolineato", false, false, true); p.Append(r);
                r = OpenXmlTools.CreaRun("Testo grassetto, corsivo, sottolineato, colorato", true, true, true, "993300"); p.Append(r);
                docBody.Append(p);

                p = OpenXmlTools.CreaParagrafo();
                r = OpenXmlTools.CreaRun("Testo con font arial 34", false, false, false, "000000", "Arial", 34);
                p.Append(r);
                docBody.Append(p);
            }
        }

    }
}
