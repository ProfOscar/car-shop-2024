using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JustificationValues = DocumentFormat.OpenXml.Wordprocessing.JustificationValues;

namespace CarShop_Library
{
    public class OpenXmlTools
    {
        public static WordprocessingDocument CreaDocumento(string filePath)
        {
            WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document, true);

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

            return wordDocument;
        }

        public static Style CreaStile(WordprocessingDocument wordprocessingDocument,
            string styleName, string giustificazione = "left",
            string colore = "000000", string fontFace = "Calibri", double fontSize = 11,
            int spazioPrima = 0, int spazioDopo = 8)
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
            JustificationValues justificationValues = getJustificationValues(giustificazione);
            paragraphProperties.Justification = new Justification() { Val = justificationValues };
            string before = (spazioPrima * 20).ToString();
            string after = (spazioDopo * 20).ToString();
            paragraphProperties.SpacingBetweenLines = new SpacingBetweenLines() { Before = before, After = after };
            style.Append(paragraphProperties);

            StyleRunProperties runProperties = new StyleRunProperties();
            runProperties.Color = new Color() { Val = colore };
            runProperties.RunFonts = new RunFonts() { Ascii = fontFace };
            runProperties.FontSize = new FontSize() { Val = (fontSize * 2).ToString() };
            style.Append(runProperties);

            wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.Append(style);
            wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.Save();

            return style;
        }

        public static Paragraph CreaParagrafoConStile(string contenuto, string nomeStile)
        {
            Paragraph p = new Paragraph();
            ParagraphProperties pp = new ParagraphProperties();
            pp.ParagraphStyleId = new ParagraphStyleId() { Val = nomeStile };
            p.Append(pp);
            if (contenuto.Length > 0)
            {
                Run r = new Run();
                Text t = new Text(contenuto);
                r.Append(t);
                p.Append(r);
            }
            return p;
        }

        public static Paragraph CreaParagrafo(string contenuto = "", string giustificazione = "left",
            bool isGrassetto = false, bool isCorsivo = false, bool isSottolineato = false,
            string colore = "000000", string fontFace = "Calibri", double fontSize = 11)
        {
            Paragraph paragraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            JustificationValues justificationValues = getJustificationValues(giustificazione);
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

        #region Metodi privati
        private static JustificationValues getJustificationValues(string giustificazione)
        {
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
            return justificationValues;
        }
        #endregion

    }
}
