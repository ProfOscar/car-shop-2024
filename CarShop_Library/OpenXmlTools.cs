﻿using DocumentFormat.OpenXml;
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

                // Creo lo StyleDefinitionPart per supportare la creazione di stili personalizzati
                StyleDefinitionsPart styleDefinitionsPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                styleDefinitionsPart.Styles = new Styles();
                styleDefinitionsPart.Styles.Save();

                // Creo il documento vero e proprio
                mainPart.Document = new Document();
                Body docBody = new Body();
                mainPart.Document.Append(docBody);

                // Aggiungo la stringa per i paragrafi di test
                string lorem = @"Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed aliquet mauris in magna finibus, ut porttitor felis condimentum. Cras sed hendrerit ex. Sed porta dictum purus eu dictum. Donec hendrerit aliquet mollis. Maecenas volutpat lacus eu lorem porta, quis imperdiet nibh pharetra. Sed ac eros diam. Sed ex libero, commodo in iaculis nec, scelerisque in erat. Proin ultricies hendrerit volutpat. Vivamus porttitor, nibh in maximus gravida, enim arcu porta leo, ac porttitor enim elit vel sapien.";

                // paragrafi con style
                Style style1 = CreaStile("Codice", "left", "CCCCCC", "Courier New", 10);
                mainPart.StyleDefinitionsPart.Styles.Append(style1);
                mainPart.StyleDefinitionsPart.Styles.Save();
                // occorrerà creare un metodo CreaParagrafoConStile(...)
                Paragraph p1 = new Paragraph();
                ParagraphProperties pp1 = new ParagraphProperties();
                pp1.ParagraphStyleId = new ParagraphStyleId() { Val = "Codice" };
                p1.Append(pp1);
                // Run Heading1
                Run rH1 = new Run();
                Text tH1 = new Text(lorem);
                rH1.Append(tH1);
                p1.Append(rH1);
                // Add your heading to docx body
                docBody.Append(p1);


                // 3 paragrafi semplici con diversa giustificazione
                docBody.Append(CreaParagrafo(lorem));
                docBody.Append(CreaParagrafo(lorem, "center"));
                docBody.Append(CreaParagrafo(lorem, "right"));
                docBody.Append(CreaParagrafo(lorem, "distribute"));

                // 1 paragrafo formattato in modo omogeneo
                docBody.Append(CreaParagrafo(lorem, "left", false, true, false, "77FF33", "Tahoma", 15));

                // un paragrafo con il contenuto formattato nei diversi run
                Paragraph p = CreaParagrafo("", "center");
                Run r = CreaRun("Testo normale"); p.Append(r);
                r = CreaRun("Testo grassetto", true); p.Append(r);
                r = CreaRun("Testo corsivo", false, true); p.Append(r);
                r = CreaRun("Testo sottolineato", false, false, true); p.Append(r);
                r = CreaRun("Testo grassetto, corsivo, sottolineato, colorato", true, true, true, "993300"); p.Append(r);
                docBody.Append(p);

                p = CreaParagrafo();
                r = CreaRun("Testo con font arial 34", false, false, false, "000000", "Arial", 34);
                p.Append(r);
                docBody.Append(p);
            }
        }

        public static Paragraph CreaParagrafo(string contenuto = "", string giustificazione = "left",
            bool isGrassetto = false, bool isCorsivo = false, bool isSottolineato = false,
            string colore = "000000", string fontFace = "Calibri", double fontSize = 11)
        {
            Paragraph paragraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            JustificationValues justificationValues;
            switch (giustificazione)
            {
                case "left":
                    justificationValues = JustificationValues.Left; break;
                case "center":
                    justificationValues = JustificationValues.Center; break;
                case "right":
                    justificationValues = JustificationValues.Right; break;
                case "distribute":
                    justificationValues = JustificationValues.Distribute; break;
                default:
                    break;
            }
            paragraphProperties.Justification = new Justification() { Val = justificationValues };
            paragraph.Append(paragraphProperties);

            if (contenuto != "") paragraph.Append(CreaRun(contenuto, 
                isGrassetto, isCorsivo, isSottolineato,
                colore, fontFace, fontSize));
            return paragraph;
        }

        public static Run CreaRun(string contenuto,
            bool isGrassetto = false, bool isCorsivo = false, bool isSottolineato = false,
            string colore = "000000", string fontFace = "Calibri", double fontSize = 11)
        {
            Run run = new Run();
            RunProperties runProperties = new RunProperties();

            if (isGrassetto) runProperties.Bold = new Bold();
            if (isCorsivo) runProperties.Italic = new Italic();
            if (isSottolineato) runProperties.Underline = new Underline() { Val = UnderlineValues.Single };
            runProperties.Color = new Color() { Val = colore };
            runProperties.RunFonts = new RunFonts() { Ascii = fontFace };
            runProperties.FontSize = new FontSize() { Val = (fontSize * 2).ToString() };
            run.Append(runProperties);

            Text text = new Text(contenuto);
            run.Append(text);

            return run;
        }

        public static Style CreaStile(string styleName, string giustificazione = "left",
            string colore = "000000", string fontFace = "Calibri", double fontSize = 11)
        {
            Style style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = styleName,
                CustomStyle = true,
                StyleName = new StyleName() { Val = styleName },
                BasedOn = new BasedOn() { Val = "Normal" },
                NextParagraphStyle = new NextParagraphStyle() { Val = "Normal" },
                UIPriority = new UIPriority() { Val = 900 },
                Default = false
            };

            StyleParagraphProperties paragraphProperties = new StyleParagraphProperties();
            JustificationValues justificationValues;
            switch (giustificazione)
            {
                case "left":
                    justificationValues = JustificationValues.Left; break;
                case "center":
                    justificationValues = JustificationValues.Center; break;
                case "right":
                    justificationValues = JustificationValues.Right; break;
                case "distribute":
                    justificationValues = JustificationValues.Distribute; break;
                default:
                    break;
            }
            paragraphProperties.Justification = new Justification() { Val = justificationValues };
            style.Append(paragraphProperties);

            StyleRunProperties runProperties = new StyleRunProperties();
            runProperties.Color = new Color() { Val = colore };
            runProperties.RunFonts = new RunFonts() { Ascii = fontFace };
            runProperties.FontSize = new FontSize() { Val = (fontSize * 2).ToString() };
            style.Append(runProperties);

            return style;
        }
    }
}
