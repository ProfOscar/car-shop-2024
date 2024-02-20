using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CarShop_Library
{
    public class OpenXmlTools
    {
        public static void GeneraVolantinoDocx(string filePath)
        {
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document, true))
            {
                // Creo il MainDocumentPart che è indispensabile
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Creo il documento vero e proprio
                mainPart.Document = new Document();
                Body docBody = new Body();
                mainPart.Document.Append(docBody);

                // Aggiungo un paragrafo di test
                string lorem = @"Lorem ipsum dolor sit amet, consectetur adipiscing elit. 
                    Sed aliquet mauris in magna finibus, ut porttitor felis condimentum. 
                    Cras sed hendrerit ex. Sed porta dictum purus eu dictum. 
                    Donec hendrerit aliquet mollis. Maecenas volutpat lacus eu lorem porta, 
                    quis imperdiet nibh pharetra. Sed ac eros diam. Sed ex libero, commodo in 
                    iaculis nec, scelerisque in erat. Proin ultricies hendrerit volutpat. 
                    Vivamus porttitor, nibh in maximus gravida, enim arcu porta leo, ac porttitor 
                    enim elit vel sapien.";

                docBody.Append(CreaParagrafo(lorem, JustificationValues.Left));
                docBody.Append(CreaParagrafo(lorem, JustificationValues.Center));
                docBody.Append(CreaParagrafo(lorem, JustificationValues.Right));
            }
        }

        public static Paragraph CreaParagrafo(string contenuto,
            JustificationValues giustificazione)
        {
            Paragraph paragraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            paragraphProperties.Justification = new Justification() { Val = giustificazione };
            paragraph.Append(paragraphProperties);

            paragraph.Append(CreaRun(contenuto));
            return paragraph;
        }

        public static Run CreaRun(string contenuto,
            bool isGrassetto = false, bool isCorsivo = false, bool isSottolineato = false,
            string colore = "000000")
        {
            Run run = new Run();
            RunProperties runProperties = new RunProperties();

            if (isGrassetto) runProperties.Bold = new Bold();
            if (isCorsivo) runProperties.Italic = new Italic();
            if (isSottolineato) runProperties.Underline = new Underline();
            runProperties.Color = new Color() { Val = colore };
            run.Append(runProperties);

            Text text = new Text(contenuto);
            run.Append(text);

            return run;
        }
    }
}
