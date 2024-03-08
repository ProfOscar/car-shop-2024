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

            // Creo il NumberingDefinitionsPart per supportare la gestione degli elenchi puntati
            NumberingDefinitionsPart numberingDefinitionsPart = mainPart.AddNewPart<NumberingDefinitionsPart>("UnorderedList");
            Numbering numbering = new Numbering(
                new AbstractNum(
                    new Level(
                        new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        new LevelText() { Val = "\u2022" } // 2022 è il codice unicode del punto elenco di word
                    ) { LevelIndex = 0 }
                ) { AbstractNumberId = 1 },
                new NumberingInstance(new AbstractNumId() { Val = 1 }) { NumberID = 1 }
            );
            numbering.Save(numberingDefinitionsPart);

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

        public static Table CreaTabella(string[,] contenuto,
            string giustificazioneTabella = "left", string giustificazioneCelle = "left",
            string coloreBordi = "333333", string coloreTesto = "333333",
            int margine = 80)
        {
            Table table = new Table();
            ModificaProprietaTabella(table, giustificazioneTabella, coloreBordi, margine);
            for (int i = 0; i < contenuto.GetLength(0); i++)
            {
                TableRow tableRow = new TableRow();
                for (int j = 0; j < contenuto.GetLength(1); j++)
                {
                    TableCell tableCell = new TableCell();
                    Paragraph paragraph = CreaParagrafo(contenuto[i, j], giustificazioneCelle, false, false, false, coloreTesto);
                    tableCell.Append(paragraph);
                    tableRow.Append(tableCell);
                }
                table.Append(tableRow);
            }
            return table;
        }

        public static List<Paragraph> CreaElenco(string[] contenuto, bool isOrdered = false,
            string colore = "000000", string fontFace = "Calibri", double fontSize = 11)
        {
            List<Paragraph> list = new List<Paragraph>();

            // list properties
            SpacingBetweenLines spacingBetweenLines = new SpacingBetweenLines() { After = "0" };
            Indentation indentation = new Indentation() { Left = "260", Hanging = "240" };
            NumberingProperties numberingProperties = new NumberingProperties(
                    new NumberingLevelReference() { Val = 0 },
                    new NumberingId() { Val = isOrdered ? 2 : 1 }
                );
            ParagraphProperties paragraphProperties = new ParagraphProperties(
                spacingBetweenLines, indentation, numberingProperties);
            paragraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = "ListParagraph" };

            // content
            for (int i = 0; i < contenuto.Length; i++)
            {
                Paragraph paragraph = CreaParagrafo(contenuto[i], 
                    default, default, default, default, 
                    colore, fontFace, fontSize);
                paragraph.ParagraphProperties = new ParagraphProperties(paragraphProperties.OuterXml);
                list.Add(paragraph);
            }

            return list;
        }
        #region Metodi privati

        private static void ModificaProprietaTabella(Table table, string giustificazione, string coloreBordi, int margine)
        {
            // table justification 
            TableProperties tableProperties = new TableProperties();
            TableJustification tableJustification = getTableJustification(giustificazione);
            tableProperties.Append(tableJustification);

            // table borders
            TableBorders tableBorders = new TableBorders();
            TopBorder topBorder = new TopBorder() { Val = BorderValues.Thick, Color = coloreBordi };
            tableBorders.Append(topBorder);
            BottomBorder bottomBorder = new BottomBorder() { Val = BorderValues.Thick, Color = coloreBordi };
            tableBorders.Append(bottomBorder);
            RightBorder rightBorder = new RightBorder() { Val = BorderValues.Thick, Color = coloreBordi };
            tableBorders.Append(rightBorder);
            LeftBorder leftBorder = new LeftBorder() { Val = BorderValues.Thick, Color = coloreBordi };
            tableBorders.Append(leftBorder);
            InsideHorizontalBorder insideHorizontalBorder = new InsideHorizontalBorder() { Val = BorderValues.Thick, Color = coloreBordi };
            tableBorders.Append(insideHorizontalBorder);
            InsideVerticalBorder insideVerticalBorder = new InsideVerticalBorder() { Val = BorderValues.Thick, Color = coloreBordi };
            tableBorders.Append(insideVerticalBorder);
            tableProperties.Append(tableBorders);

            // table margins
            TableCellMarginDefault tableCellMarginDefault = new TableCellMarginDefault(
                    new TopMargin() { Width = margine.ToString(), Type = TableWidthUnitValues.Dxa },
                    new StartMargin() { Width = margine.ToString(), Type = TableWidthUnitValues.Dxa },
                    new BottomMargin() { Width = margine.ToString(), Type = TableWidthUnitValues.Dxa },
                    new EndMargin() { Width = margine.ToString(), Type = TableWidthUnitValues.Dxa }
                );
            tableProperties.Append(tableCellMarginDefault);

            table.Append(tableProperties);
        }

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

        private static TableJustification getTableJustification(string giustificazione)
        {
            TableJustification justificationValues;
            switch (giustificazione)
            {
                case "left":
                    justificationValues = new TableJustification() { Val = TableRowAlignmentValues.Left};
                    break;
                case "center":
                    justificationValues = new TableJustification() { Val = TableRowAlignmentValues.Center };
                    break;
                case "right":
                    justificationValues = new TableJustification() { Val = TableRowAlignmentValues.Right };
                    break;
                default:
                    justificationValues = new TableJustification();
                    break;
            }
            return justificationValues;
        }
        #endregion

    }
}
