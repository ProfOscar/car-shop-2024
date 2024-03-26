using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Excel = DocumentFormat.OpenXml.Spreadsheet;

using CarShop_Library;

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
                    case 'x':
                    case 'X':
                        EsportaXlsx();
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

        private static void EsportaXlsx()
        {
            string pgmDir = AppDomain.CurrentDomain.BaseDirectory;
            string now = DateTime.Now.ToShortDateString().Replace('/', '-');
            string filePath = $"{pgmDir}/report_{now}.xlsx";
            TestReportXlsx(filePath);
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
            Console.WriteLine("X - Crea Report XLSX");
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
            using (WordprocessingDocument wordDocument = OpenXmlWordTools.CreaDocumento(filePath))
            {
                // prendo un riferimento al body del documento
                Body docBody = wordDocument.MainDocumentPart.Document.Body;

                // Aggiungo la stringa per i paragrafi di test
                string lorem = @"Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed aliquet mauris in magna finibus, ut porttitor felis condimentum. Cras sed hendrerit ex. Sed porta dictum purus eu dictum. Donec hendrerit aliquet mollis. Maecenas volutpat lacus eu lorem porta, quis imperdiet nibh pharetra. Sed ac eros diam. Sed ex libero, commodo in iaculis nec, scelerisque in erat. Proin ultricies hendrerit volutpat. Vivamus porttitor, nibh in maximus gravida, enim arcu porta leo, ac porttitor enim elit vel sapien.";

                // definisco stili
                Style titolo1Style = OpenXmlWordTools.CreaStile(wordDocument, "Mio Titolo 1", "center", "1100CC", "Comics Sans", 24, 80, 120);
                Style titolo2Style = OpenXmlWordTools.CreaStile(wordDocument, "Mio Titolo 2", "center", "AA5522", "Tahoma", 18, 60, 40);
                Style codiceStyle = OpenXmlWordTools.CreaStile(wordDocument, "Codice", "left", "CCCCCC", "Courier New", 10, 50, 100);

                // utilizzo stili
                docBody.Append(OpenXmlWordTools.CreaParagrafoConStile("Prova Mio Titolo 1", titolo1Style.StyleId));
                docBody.Append(OpenXmlWordTools.CreaParagrafoConStile("Prova Mio Titolo 2", titolo2Style.StyleId));
                docBody.Append(OpenXmlWordTools.CreaParagrafoConStile(lorem, codiceStyle.StyleId));

                // 3 paragrafi semplici con diversa giustificazione
                docBody.Append(OpenXmlWordTools.CreaParagrafo(lorem));
                docBody.Append(OpenXmlWordTools.CreaParagrafo(lorem, "center"));
                docBody.Append(OpenXmlWordTools.CreaParagrafo(lorem, "right"));
                docBody.Append(OpenXmlWordTools.CreaParagrafo(lorem, "distribute"));

                // 1 paragrafo formattato in modo omogeneo
                docBody.Append(OpenXmlWordTools.CreaParagrafo(lorem, "left", false, true, false, "77FF33", "Tahoma", 15));

                // un paragrafo con il contenuto formattato nei diversi run
                Paragraph p = OpenXmlWordTools.CreaParagrafo("", "center");
                Run r = OpenXmlWordTools.CreaRun("Testo normale"); p.Append(r);
                r = OpenXmlWordTools.CreaRun("Testo grassetto", true); p.Append(r);
                r = OpenXmlWordTools.CreaRun("Testo corsivo", false, true); p.Append(r);
                r = OpenXmlWordTools.CreaRun("Testo sottolineato", false, false, true); p.Append(r);
                r = OpenXmlWordTools.CreaRun("Testo grassetto, corsivo, sottolineato, colorato", true, true, true, "993300"); p.Append(r);
                docBody.Append(p);

                // paragrafo con font impostato nel run
                p = OpenXmlWordTools.CreaParagrafo();
                r = OpenXmlWordTools.CreaRun("Testo con font arial 34", false, false, false, "000000", "Arial", 34);
                p.Append(r);
                docBody.Append(p);

                // test hyperlink
                Paragraph pHyperlink = OpenXmlWordTools.CreaParagrafo();
                // Hyperlink hyperlink = OpenXmlTools.CreaHyperlink(wordDocument, "http://www.vallauri.edu", "Vai al sito del Vallauri");
                Hyperlink hyperlink = OpenXmlWordTools.CreaHyperlink(wordDocument,
                    "http://www.vallauri.edu", "Vai al sito del Vallauri",
                    false, false, true, "0000FF", "Tahoma", 32.5);
                pHyperlink.Append(hyperlink);
                docBody.Append(pHyperlink);

                // test hyperlink in paragrafo allineato a destra
                pHyperlink = OpenXmlWordTools.CreaParagrafo("", "right");
                hyperlink = OpenXmlWordTools.CreaHyperlink(wordDocument, "http://www.vallauri.edu", "Vai al sito del Vallauri (allineato a destra)");
                pHyperlink.Append(hyperlink);
                docBody.Append(pHyperlink);

                // test elenchi
                string[] contenutoElenchi = { "BMW Serie 3", "Jeep Compass", "Mercedes CLA", "Fiat Panda" };
                // elenco numerato
                List<Paragraph> elenco = OpenXmlWordTools.CreaElenco(contenutoElenchi, true);
                foreach (var item in elenco) docBody.Append(item);
                // elenco puntato
                elenco = OpenXmlWordTools.CreaElenco(contenutoElenchi, false);
                foreach (var item in elenco) docBody.Append(item);

                // test tabella
                string[,] contenutoTabella = {
                    { "MARCA", "MODELLO", "TARGA", "PREZZO" },
                    { "BMW", "iX2", "GG528YT", "€ 57.800" },
                    { "Jeep", "Compass", "FR508HD", "€ 35750" }
                };
                Table table = OpenXmlWordTools.CreaTabella(contenutoTabella, "center", "right", "red", "green", 380);
                docBody.Append(table);

                // test immagine
                string imageUrl = "https://www.robinsonpetshop.it/news/cms2017/wp-content/uploads/2022/07/GattinoPrimiMesi.jpg";
                Paragraph pImage = OpenXmlWordTools.AggiungiImmagine(wordDocument, imageUrl, "center", 100, 100);
                docBody.Append(pImage);
                imageUrl = "https://png.pngtree.com/png-clipart/20230507/ourmid/pngtree-tiger-walking-wildlife-scene-transparent-background-png-image_7088126.png";
                pImage = OpenXmlWordTools.AggiungiImmagine(wordDocument, imageUrl, "right");
                docBody.Append(pImage);
            }
        }

        public static void TestReportXlsx(string filePath)
        {
            string[] titolo = { "Prova Utilizzo OpenXmlExcelTools" };
            string[,] datiDiTest =
            {
                { "Articolo", "Giacenza", "Prezzo" },
                { "Cioccolata", "25.80", "12" },
                { "Caffè", "148", "22" },
                { "Pasta", "500", "7.50" }
            };
            SpreadsheetDocument reportDocument = OpenXmlExcelTools.CreaDocumento(filePath);
            using (reportDocument)
            {
                Excel.SheetData sheetData = OpenXmlExcelTools.getFirstSheetData(reportDocument);
                // riga di titolo
                Excel.Row row = OpenXmlExcelTools.CreaRiga(titolo, true);
                sheetData.Append(row);
                // aggiungo una riga vuota
                row = OpenXmlExcelTools.CreaRiga(); sheetData.Append(row);
                // riga dell'header
                row = OpenXmlExcelTools.CreaRiga();
                for (int j = 0; j < datiDiTest.GetLength(1); j++)
                    row.Append(OpenXmlExcelTools.CreaCella(datiDiTest[0, j], true, true));
                sheetData.Append(row);
                // righe dei dati   
                for (int i = 1; i < datiDiTest.GetLength(0); i++)
                {
                    row = OpenXmlExcelTools.CreaRiga();
                    for (int j = 0; j < datiDiTest.GetLength(1); j++)
                        row.Append(OpenXmlExcelTools.CreaCella(datiDiTest[i, j], false, true));
                    sheetData.Append(row);
                }      
            }
        }
    }
}
