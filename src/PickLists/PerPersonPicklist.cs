using BrickAtHeart.LUGTools.PicklistGenerator.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Options;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using BottomBorder = DocumentFormat.OpenXml.Wordprocessing.BottomBorder;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Fonts = DocumentFormat.OpenXml.Wordprocessing.Fonts;
using GridColumn = DocumentFormat.OpenXml.Wordprocessing.GridColumn;
using LeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder;
using Outline = DocumentFormat.OpenXml.Wordprocessing.Outline;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableCellBorders = DocumentFormat.OpenXml.Wordprocessing.TableCellBorders;
using TableCellProperties = DocumentFormat.OpenXml.Wordprocessing.TableCellProperties;
using TableGrid = DocumentFormat.OpenXml.Wordprocessing.TableGrid;
using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using TopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder;
using Underline = DocumentFormat.OpenXml.Wordprocessing.Underline;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace BrickAtHeart.LUGTools.PicklistGenerator
{
    internal class PerPersonPicklist
    {
        public PerPersonPicklist(IOptions<PicklistGeneratorOptions> options)
        {
            this.options = options.Value;
        }

        internal void Generate(List<Order> orders)
        {
            using WordprocessingDocument package = WordprocessingDocument.Create(options.PerPersonFileName, WordprocessingDocumentType.Document);

            ExtendedFilePropertiesPart extendedFilePropertiesPart = package.AddNewPart<ExtendedFilePropertiesPart>("rId2");
            GenerateExtendedFilePropertiesPartContent(extendedFilePropertiesPart);

            MainDocumentPart mainDocumentPart = package.AddMainDocumentPart();
            GenerateMainDocumentPartContent(mainDocumentPart, orders);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart.AddNewPart<StyleDefinitionsPart>("rId1");
            GenerateStyleDefinitionsPartContent(styleDefinitionsPart1);

            FontTablePart fontTablePart1 = mainDocumentPart.AddNewPart<FontTablePart>("rId4");
            GenerateFontTablePartContent(fontTablePart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart.AddNewPart<DocumentSettingsPart>("rId5");
            GenerateDocumentSettingsPartContent(documentSettingsPart1);

            SetPackageProperties(package);
        }

        private void GenerateExtendedFilePropertiesPartContent(ExtendedFilePropertiesPart extendedFilePropertiesPart)
        {
            Ap.Properties properties = new Ap.Properties();
            properties.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application = new Ap.Application {Text = "LUGTools/Picklists"};

            properties.Append(application);

            extendedFilePropertiesPart.Properties = properties;
        }

        private void GenerateMainDocumentPartContent(MainDocumentPart mainDocumentPart, List<Order> orders)
        {
            Document document = new Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            document.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
            document.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            document.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            document.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");

            Body body = new Body();

            IEnumerable<IGrouping<string, Order>> perPartGroupingList = orders.GroupBy(o => o.Part.LegoElementId);

            foreach (IGrouping<string, Order> grouping in perPartGroupingList)
            {
                ImagePart imagePart = mainDocumentPart.AddNewPart<ImagePart>("image/png", $"rId{grouping.Key}");
                GenerateImagePartContent(imagePart, grouping.Key);
            }

            IEnumerable<IGrouping<string, Order>> groupingList = orders.GroupBy(o => o.Person.FullName).OrderBy(o => o.First().Person.Index);

            foreach (IGrouping<string, Order> grouping in groupingList)
            {
                Person person = grouping.FirstOrDefault()?.Person;

                if (person == null)
                {
                    continue;
                }

                Paragraph paragraph1 = new Paragraph();

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Normal" };
                BiDi biDi1 = new BiDi() { Val = false };
                Justification justification1 = new Justification() { Val = JustificationValues.Left };

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                FontSize fontSize1 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "32" };

                paragraphMarkRunProperties1.Append(fontSize1);
                paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

                paragraphProperties1.Append(paragraphStyleId1);
                paragraphProperties1.Append(biDi1);
                paragraphProperties1.Append(justification1);
                paragraphProperties1.Append(paragraphMarkRunProperties1);

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                FontSize fontSize2 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "32" };

                runProperties1.Append(fontSize2);
                runProperties1.Append(fontSizeComplexScript2);
                Text text1 = new Text {Text = $"{person.Index}. {person.FullName}"};

                run1.Append(runProperties1);
                run1.Append(text1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                Paragraph paragraph2 = new Paragraph();

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Normal" };
                BiDi biDi2 = new BiDi() { Val = false };
                Justification justification2 = new Justification() { Val = JustificationValues.Left };

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                FontSize fontSize3 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "32" };

                paragraphMarkRunProperties2.Append(fontSize3);
                paragraphMarkRunProperties2.Append(fontSizeComplexScript3);

                paragraphProperties2.Append(paragraphStyleId2);
                paragraphProperties2.Append(biDi2);
                paragraphProperties2.Append(justification2);
                paragraphProperties2.Append(paragraphMarkRunProperties2);

                Run run2 = new Run();

                RunProperties runProperties2 = new RunProperties();
                FontSize fontSize4 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "32" };

                runProperties2.Append(fontSize4);
                runProperties2.Append(fontSizeComplexScript4);

                run2.Append(runProperties2);

                paragraph2.Append(paragraphProperties2);
                paragraph2.Append(run2);

                body.Append(paragraph1);
                body.Append(paragraph2);

                Table table1 = new Table();

                TableProperties tableProperties1 = new TableProperties();
                TableWidth tableWidth1 = new TableWidth() { Width = "9972", Type = TableWidthUnitValues.Dxa };
                TableJustification tableJustification1 = new TableJustification() { Val = TableRowAlignmentValues.Left };
                TableIndentation tableIndentation1 = new TableIndentation() { Width = 55, Type = TableWidthUnitValues.Dxa };

                TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
                TopMargin topMargin1 = new TopMargin() { Width = "55", Type = TableWidthUnitValues.Dxa };
                TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 55, Type = TableWidthValues.Dxa };
                BottomMargin bottomMargin1 = new BottomMargin() { Width = "55", Type = TableWidthUnitValues.Dxa };
                TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 55, Type = TableWidthValues.Dxa };

                tableCellMarginDefault1.Append(topMargin1);
                tableCellMarginDefault1.Append(tableCellLeftMargin1);
                tableCellMarginDefault1.Append(bottomMargin1);
                tableCellMarginDefault1.Append(tableCellRightMargin1);

                tableProperties1.Append(tableWidth1);
                tableProperties1.Append(tableJustification1);
                tableProperties1.Append(tableIndentation1);
                tableProperties1.Append(tableCellMarginDefault1);

                TableGrid tableGrid1 = new TableGrid();
                GridColumn gridColumn1 = new GridColumn() { Width = "2159" };
                GridColumn gridColumn2 = new GridColumn() { Width = "4488" };
                GridColumn gridColumn3 = new GridColumn() { Width = "3325" };

                tableGrid1.Append(gridColumn1);
                tableGrid1.Append(gridColumn2);
                tableGrid1.Append(gridColumn3);

                TableRow tableRow1 = new TableRow();
                TableRowProperties tableRowProperties1 = new TableRowProperties();

                TableCell tableCell1 = new TableCell();

                TableCellProperties tableCellProperties1 = new TableCellProperties();
                TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "2159", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders1 = new TableCellBorders();
                TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };
                LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };
                BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };

                tableCellBorders1.Append(topBorder1);
                tableCellBorders1.Append(leftBorder1);
                tableCellBorders1.Append(bottomBorder1);
                Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

                tableCellProperties1.Append(tableCellWidth1);
                tableCellProperties1.Append(tableCellBorders1);
                tableCellProperties1.Append(shading1);

                Paragraph paragraph3 = new Paragraph();

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "TableContents" };
                BiDi biDi3 = new BiDi() { Val = false };
                Justification justification3 = new Justification() { Val = JustificationValues.Left };

                ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                RunFonts runFonts1 = new RunFonts() { Ascii = "Liberation Serif", HighAnsi = "Liberation Serif" };
                Bold bold1 = new Bold() { Val = false };
                Bold bold2 = new Bold() { Val = false };
                BoldComplexScript boldComplexScript1 = new BoldComplexScript() { Val = false };
                Italic italic1 = new Italic() { Val = false };
                Italic italic2 = new Italic() { Val = false };
                ItalicComplexScript italicComplexScript1 = new ItalicComplexScript() { Val = false };
                Strike strike1 = new Strike() { Val = false };
                DoubleStrike doubleStrike1 = new DoubleStrike() { Val = false };
                Outline outline1 = new Outline() { Val = false };
                Shadow shadow1 = new Shadow() { Val = false };
                Color color1 = new Color() { Val = "000000" };
                FontSize fontSize5 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "24" };
                Underline underline1 = new Underline() { Val = UnderlineValues.None };

                paragraphMarkRunProperties3.Append(runFonts1);
                paragraphMarkRunProperties3.Append(bold1);
                paragraphMarkRunProperties3.Append(bold2);
                paragraphMarkRunProperties3.Append(boldComplexScript1);
                paragraphMarkRunProperties3.Append(italic1);
                paragraphMarkRunProperties3.Append(italic2);
                paragraphMarkRunProperties3.Append(italicComplexScript1);
                paragraphMarkRunProperties3.Append(strike1);
                paragraphMarkRunProperties3.Append(doubleStrike1);
                paragraphMarkRunProperties3.Append(outline1);
                paragraphMarkRunProperties3.Append(shadow1);
                paragraphMarkRunProperties3.Append(color1);
                paragraphMarkRunProperties3.Append(fontSize5);
                paragraphMarkRunProperties3.Append(fontSizeComplexScript5);
                paragraphMarkRunProperties3.Append(underline1);

                paragraphProperties3.Append(paragraphStyleId3);
                paragraphProperties3.Append(biDi3);
                paragraphProperties3.Append(justification3);
                paragraphProperties3.Append(paragraphMarkRunProperties3);

                Run run3 = new Run();

                RunProperties runProperties3 = new RunProperties();
                Bold bold3 = new Bold() { Val = false };
                BoldComplexScript boldComplexScript2 = new BoldComplexScript() { Val = false };
                Italic italic3 = new Italic() { Val = false };
                ItalicComplexScript italicComplexScript2 = new ItalicComplexScript() { Val = false };
                Strike strike2 = new Strike() { Val = false };
                DoubleStrike doubleStrike2 = new DoubleStrike() { Val = false };
                Outline outline2 = new Outline() { Val = false };
                Shadow shadow2 = new Shadow() { Val = false };
                Color color2 = new Color() { Val = "000000" };
                FontSize fontSize6 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "24" };
                Underline underline2 = new Underline() { Val = UnderlineValues.None };

                runProperties3.Append(bold3);
                runProperties3.Append(boldComplexScript2);
                runProperties3.Append(italic3);
                runProperties3.Append(italicComplexScript2);
                runProperties3.Append(strike2);
                runProperties3.Append(doubleStrike2);
                runProperties3.Append(outline2);
                runProperties3.Append(shadow2);
                runProperties3.Append(color2);
                runProperties3.Append(fontSize6);
                runProperties3.Append(fontSizeComplexScript6);
                runProperties3.Append(underline2);
                Text text2 = new Text {Text = "Image"};

                run3.Append(runProperties3);
                run3.Append(text2);

                paragraph3.Append(paragraphProperties3);
                paragraph3.Append(run3);

                tableCell1.Append(tableCellProperties1);
                tableCell1.Append(paragraph3);

                TableCell tableCell2 = new TableCell();

                TableCellProperties tableCellProperties2 = new TableCellProperties();
                TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "4488", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders2 = new TableCellBorders();
                TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };
                LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };
                BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };

                tableCellBorders2.Append(topBorder2);
                tableCellBorders2.Append(leftBorder2);
                tableCellBorders2.Append(bottomBorder2);
                Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

                tableCellProperties2.Append(tableCellWidth2);
                tableCellProperties2.Append(tableCellBorders2);
                tableCellProperties2.Append(shading2);

                Paragraph paragraph4 = new Paragraph();

                ParagraphProperties paragraphProperties4 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "TableContents" };
                BiDi biDi4 = new BiDi() { Val = false };
                Justification justification4 = new Justification() { Val = JustificationValues.Left };

                ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
                RunFonts runFonts2 = new RunFonts() { Ascii = "Liberation Serif", HighAnsi = "Liberation Serif" };
                Bold bold4 = new Bold() { Val = false };
                Bold bold5 = new Bold() { Val = false };
                BoldComplexScript boldComplexScript3 = new BoldComplexScript() { Val = false };
                Italic italic4 = new Italic() { Val = false };
                Italic italic5 = new Italic() { Val = false };
                ItalicComplexScript italicComplexScript3 = new ItalicComplexScript() { Val = false };
                Strike strike3 = new Strike() { Val = false };
                DoubleStrike doubleStrike3 = new DoubleStrike() { Val = false };
                Outline outline3 = new Outline() { Val = false };
                Shadow shadow3 = new Shadow() { Val = false };
                Color color3 = new Color() { Val = "000000" };
                FontSize fontSize7 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "24" };
                Underline underline3 = new Underline() { Val = UnderlineValues.None };

                paragraphMarkRunProperties4.Append(runFonts2);
                paragraphMarkRunProperties4.Append(bold4);
                paragraphMarkRunProperties4.Append(bold5);
                paragraphMarkRunProperties4.Append(boldComplexScript3);
                paragraphMarkRunProperties4.Append(italic4);
                paragraphMarkRunProperties4.Append(italic5);
                paragraphMarkRunProperties4.Append(italicComplexScript3);
                paragraphMarkRunProperties4.Append(strike3);
                paragraphMarkRunProperties4.Append(doubleStrike3);
                paragraphMarkRunProperties4.Append(outline3);
                paragraphMarkRunProperties4.Append(shadow3);
                paragraphMarkRunProperties4.Append(color3);
                paragraphMarkRunProperties4.Append(fontSize7);
                paragraphMarkRunProperties4.Append(fontSizeComplexScript7);
                paragraphMarkRunProperties4.Append(underline3);

                paragraphProperties4.Append(paragraphStyleId4);
                paragraphProperties4.Append(biDi4);
                paragraphProperties4.Append(justification4);
                paragraphProperties4.Append(paragraphMarkRunProperties4);

                Run run4 = new Run();

                RunProperties runProperties4 = new RunProperties();
                Bold bold6 = new Bold() { Val = false };
                BoldComplexScript boldComplexScript4 = new BoldComplexScript() { Val = false };
                Italic italic6 = new Italic() { Val = false };
                ItalicComplexScript italicComplexScript4 = new ItalicComplexScript() { Val = false };
                Strike strike4 = new Strike() { Val = false };
                DoubleStrike doubleStrike4 = new DoubleStrike() { Val = false };
                Outline outline4 = new Outline() { Val = false };
                Shadow shadow4 = new Shadow() { Val = false };
                Color color4 = new Color() { Val = "000000" };
                FontSize fontSize8 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "24" };
                Underline underline4 = new Underline() { Val = UnderlineValues.None };

                runProperties4.Append(bold6);
                runProperties4.Append(boldComplexScript4);
                runProperties4.Append(italic6);
                runProperties4.Append(italicComplexScript4);
                runProperties4.Append(strike4);
                runProperties4.Append(doubleStrike4);
                runProperties4.Append(outline4);
                runProperties4.Append(shadow4);
                runProperties4.Append(color4);
                runProperties4.Append(fontSize8);
                runProperties4.Append(fontSizeComplexScript8);
                runProperties4.Append(underline4);
                Text text3 = new Text {Text = "Element Name"};
                Break break1 = new Break();
                Text text4 = new Text {Text = "LEGO Color / Bricklink Color"};

                run4.Append(runProperties4);
                run4.Append(text3);
                run4.Append(break1);
                run4.Append(text4);

                paragraph4.Append(paragraphProperties4);
                paragraph4.Append(run4);

                tableCell2.Append(tableCellProperties2);
                tableCell2.Append(paragraph4);

                TableCell tableCell3 = new TableCell();

                TableCellProperties tableCellProperties3 = new TableCellProperties();
                TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "3325", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders3 = new TableCellBorders();
                TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };
                LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };
                BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };
                RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };

                tableCellBorders3.Append(topBorder3);
                tableCellBorders3.Append(leftBorder3);
                tableCellBorders3.Append(bottomBorder3);
                tableCellBorders3.Append(rightBorder1);
                Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

                tableCellProperties3.Append(tableCellWidth3);
                tableCellProperties3.Append(tableCellBorders3);
                tableCellProperties3.Append(shading3);

                Paragraph paragraph5 = new Paragraph();

                ParagraphProperties paragraphProperties5 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId5 = new ParagraphStyleId() { Val = "TableContents" };
                BiDi biDi5 = new BiDi() { Val = false };
                Justification justification5 = new Justification() { Val = JustificationValues.Left };

                ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
                RunFonts runFonts3 = new RunFonts() { Ascii = "Liberation Serif", HighAnsi = "Liberation Serif" };
                Bold bold7 = new Bold() { Val = false };
                Bold bold8 = new Bold() { Val = false };
                BoldComplexScript boldComplexScript5 = new BoldComplexScript() { Val = false };
                Italic italic7 = new Italic() { Val = false };
                Italic italic8 = new Italic() { Val = false };
                ItalicComplexScript italicComplexScript5 = new ItalicComplexScript() { Val = false };
                Strike strike5 = new Strike() { Val = false };
                DoubleStrike doubleStrike5 = new DoubleStrike() { Val = false };
                Outline outline5 = new Outline() { Val = false };
                Shadow shadow5 = new Shadow() { Val = false };
                Color color5 = new Color() { Val = "000000" };
                FontSize fontSize9 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "24" };
                Underline underline5 = new Underline() { Val = UnderlineValues.None };

                paragraphMarkRunProperties5.Append(runFonts3);
                paragraphMarkRunProperties5.Append(bold7);
                paragraphMarkRunProperties5.Append(bold8);
                paragraphMarkRunProperties5.Append(boldComplexScript5);
                paragraphMarkRunProperties5.Append(italic7);
                paragraphMarkRunProperties5.Append(italic8);
                paragraphMarkRunProperties5.Append(italicComplexScript5);
                paragraphMarkRunProperties5.Append(strike5);
                paragraphMarkRunProperties5.Append(doubleStrike5);
                paragraphMarkRunProperties5.Append(outline5);
                paragraphMarkRunProperties5.Append(shadow5);
                paragraphMarkRunProperties5.Append(color5);
                paragraphMarkRunProperties5.Append(fontSize9);
                paragraphMarkRunProperties5.Append(fontSizeComplexScript9);
                paragraphMarkRunProperties5.Append(underline5);

                paragraphProperties5.Append(paragraphStyleId5);
                paragraphProperties5.Append(biDi5);
                paragraphProperties5.Append(justification5);
                paragraphProperties5.Append(paragraphMarkRunProperties5);

                Run run5 = new Run();

                RunProperties runProperties5 = new RunProperties();
                Bold bold9 = new Bold() { Val = false };
                BoldComplexScript boldComplexScript6 = new BoldComplexScript() { Val = false };
                Italic italic9 = new Italic() { Val = false };
                ItalicComplexScript italicComplexScript6 = new ItalicComplexScript() { Val = false };
                Strike strike6 = new Strike() { Val = false };
                DoubleStrike doubleStrike6 = new DoubleStrike() { Val = false };
                Outline outline6 = new Outline() { Val = false };
                Shadow shadow6 = new Shadow() { Val = false };
                Color color6 = new Color() { Val = "000000" };
                FontSize fontSize10 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "24" };
                Underline underline6 = new Underline() { Val = UnderlineValues.None };

                runProperties5.Append(bold9);
                runProperties5.Append(boldComplexScript6);
                runProperties5.Append(italic9);
                runProperties5.Append(italicComplexScript6);
                runProperties5.Append(strike6);
                runProperties5.Append(doubleStrike6);
                runProperties5.Append(outline6);
                runProperties5.Append(shadow6);
                runProperties5.Append(color6);
                runProperties5.Append(fontSize10);
                runProperties5.Append(fontSizeComplexScript10);
                runProperties5.Append(underline6);
                Text text5 = new Text {Text = "Quantity"};

                run5.Append(runProperties5);
                run5.Append(text5);

                paragraph5.Append(paragraphProperties5);
                paragraph5.Append(run5);

                tableCell3.Append(tableCellProperties3);
                tableCell3.Append(paragraph5);

                tableRow1.Append(tableRowProperties1);
                tableRow1.Append(tableCell1);
                tableRow1.Append(tableCell2);
                tableRow1.Append(tableCell3);

                table1.Append(tableProperties1);
                table1.Append(tableGrid1);
                table1.Append(tableRow1);

                foreach (Order order in grouping)
                {
                    TableRow tableRow2 = new TableRow();

                    TableRowProperties tableRowProperties2 = new TableRowProperties();
                    CantSplit cantSplit = new CantSplit { Val = new EnumValue<OnOffOnlyValues>() { InnerText = "true" } };
                    tableRowProperties2.Append(cantSplit);

                    TableCell tableCell4 = new TableCell();

                    TableCellProperties tableCellProperties4 = new TableCellProperties();
                    TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "2159", Type = TableWidthUnitValues.Dxa };

                    TableCellBorders tableCellBorders4 = new TableCellBorders();
                    LeftBorder leftBorder4 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };
                    BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };

                    tableCellBorders4.Append(leftBorder4);
                    tableCellBorders4.Append(bottomBorder4);
                    Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

                    tableCellProperties4.Append(tableCellWidth4);
                    tableCellProperties4.Append(tableCellBorders4);
                    tableCellProperties4.Append(shading4);

                    Paragraph paragraph6 = new Paragraph();

                    ParagraphProperties paragraphProperties6 = new ParagraphProperties();
                    ParagraphStyleId paragraphStyleId6 = new ParagraphStyleId() { Val = "TableContents" };
                    BiDi biDi6 = new BiDi() { Val = false };
                    Justification justification6 = new Justification() { Val = JustificationValues.Left };

                    ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
                    RunFonts runFonts4 = new RunFonts() { Ascii = "Liberation Serif", HighAnsi = "Liberation Serif" };
                    Bold bold10 = new Bold() { Val = false };
                    Bold bold11 = new Bold() { Val = false };
                    BoldComplexScript boldComplexScript7 = new BoldComplexScript() { Val = false };
                    Italic italic10 = new Italic() { Val = false };
                    Italic italic11 = new Italic() { Val = false };
                    ItalicComplexScript italicComplexScript7 = new ItalicComplexScript() { Val = false };
                    Strike strike7 = new Strike() { Val = false };
                    DoubleStrike doubleStrike7 = new DoubleStrike() { Val = false };
                    Outline outline7 = new Outline() { Val = false };
                    Shadow shadow7 = new Shadow() { Val = false };
                    Color color7 = new Color() { Val = "000000" };
                    FontSize fontSize11 = new FontSize() { Val = "24" };
                    FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "24" };
                    Underline underline7 = new Underline() { Val = UnderlineValues.None };

                    paragraphMarkRunProperties6.Append(runFonts4);
                    paragraphMarkRunProperties6.Append(bold10);
                    paragraphMarkRunProperties6.Append(bold11);
                    paragraphMarkRunProperties6.Append(boldComplexScript7);
                    paragraphMarkRunProperties6.Append(italic10);
                    paragraphMarkRunProperties6.Append(italic11);
                    paragraphMarkRunProperties6.Append(italicComplexScript7);
                    paragraphMarkRunProperties6.Append(strike7);
                    paragraphMarkRunProperties6.Append(doubleStrike7);
                    paragraphMarkRunProperties6.Append(outline7);
                    paragraphMarkRunProperties6.Append(shadow7);
                    paragraphMarkRunProperties6.Append(color7);
                    paragraphMarkRunProperties6.Append(fontSize11);
                    paragraphMarkRunProperties6.Append(fontSizeComplexScript11);
                    paragraphMarkRunProperties6.Append(underline7);

                    paragraphProperties6.Append(paragraphStyleId6);
                    paragraphProperties6.Append(biDi6);
                    paragraphProperties6.Append(justification6);
                    paragraphProperties6.Append(paragraphMarkRunProperties6);

                    Run run6 = new Run();

                    RunProperties runProperties6 = new RunProperties();
                    Bold bold12 = new Bold() { Val = false };
                    BoldComplexScript boldComplexScript8 = new BoldComplexScript() { Val = false };
                    Italic italic12 = new Italic() { Val = false };
                    ItalicComplexScript italicComplexScript8 = new ItalicComplexScript() { Val = false };
                    Strike strike8 = new Strike() { Val = false };
                    DoubleStrike doubleStrike8 = new DoubleStrike() { Val = false };
                    Outline outline8 = new Outline() { Val = false };
                    Shadow shadow8 = new Shadow() { Val = false };
                    Color color8 = new Color() { Val = "000000" };
                    FontSize fontSize12 = new FontSize() { Val = "24" };
                    FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "24" };
                    Underline underline8 = new Underline() { Val = UnderlineValues.None };

                    runProperties6.Append(bold12);
                    runProperties6.Append(boldComplexScript8);
                    runProperties6.Append(italic12);
                    runProperties6.Append(italicComplexScript8);
                    runProperties6.Append(strike8);
                    runProperties6.Append(doubleStrike8);
                    runProperties6.Append(outline8);
                    runProperties6.Append(shadow8);
                    runProperties6.Append(color8);
                    runProperties6.Append(fontSize12);
                    runProperties6.Append(fontSizeComplexScript12);
                    runProperties6.Append(underline8);

                    Drawing drawing1 = new Drawing();

                    Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U, SimplePos = false, RelativeHeight = 2U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true };
                    Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };

                    Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Column };
                    Wp.HorizontalAlignment horizontalAlignment1 = new Wp.HorizontalAlignment {Text = "center"};

                    horizontalPosition1.Append(horizontalAlignment1);

                    Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
                    Wp.PositionOffset positionOffset1 = new Wp.PositionOffset {Text = "635"};

                    verticalPosition1.Append(positionOffset1);
                    Wp.Extent extent1 = new Wp.Extent() { Cx = 914400L, Cy = 694690L };
                    Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
                    Wp.WrapSquare wrapSquare1 = new Wp.WrapSquare() { WrapText = Wp.WrapTextValues.Largest };
                    Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = 1U, Name = $"{order.Part.LegoElementId}", Description = $"{order.Part.LegoElementDescription}" };

                    Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

                    A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoChangeAspect = true };
                    graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                    nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

                    A.Graphic graphic1 = new A.Graphic();
                    graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                    A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

                    Pic.Picture picture1 = new Pic.Picture();
                    picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

                    Pic.NonVisualPictureProperties nonVisualPictureProperties1 = new Pic.NonVisualPictureProperties();
                    Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = 1U, Name = $"{order.Part.LegoElementId}", Description = $"{order.Part.LegoElementDescription}" };

                    Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Pic.NonVisualPictureDrawingProperties();
                    A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

                    nonVisualPictureDrawingProperties1.Append(pictureLocks1);

                    nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
                    nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

                    Pic.BlipFill blipFill1 = new Pic.BlipFill();
                    A.Blip blip1 = new A.Blip { Embed = $"rId{order.Part.LegoElementId}" };

                    A.Stretch stretch1 = new A.Stretch();
                    A.FillRectangle fillRectangle1 = new A.FillRectangle();

                    stretch1.Append(fillRectangle1);

                    blipFill1.Append(blip1);
                    blipFill1.Append(stretch1);

                    Pic.ShapeProperties shapeProperties1 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                    A.Transform2D transform2D1 = new A.Transform2D();
                    A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
                    A.Extents extents1 = new A.Extents() { Cx = 914400L, Cy = 694690L };

                    transform2D1.Append(offset1);
                    transform2D1.Append(extents1);

                    A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                    A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

                    presetGeometry1.Append(adjustValueList1);

                    shapeProperties1.Append(transform2D1);
                    shapeProperties1.Append(presetGeometry1);

                    picture1.Append(nonVisualPictureProperties1);
                    picture1.Append(blipFill1);
                    picture1.Append(shapeProperties1);

                    graphicData1.Append(picture1);

                    graphic1.Append(graphicData1);

                    anchor1.Append(simplePosition1);
                    anchor1.Append(horizontalPosition1);
                    anchor1.Append(verticalPosition1);
                    anchor1.Append(extent1);
                    anchor1.Append(effectExtent1);
                    anchor1.Append(wrapSquare1);
                    anchor1.Append(docProperties1);
                    anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
                    anchor1.Append(graphic1);

                    drawing1.Append(anchor1);

                    run6.Append(runProperties6);
                    run6.Append(drawing1);

                    paragraph6.Append(paragraphProperties6);
                    paragraph6.Append(run6);

                    tableCell4.Append(tableCellProperties4);
                    tableCell4.Append(paragraph6);

                    TableCell tableCell5 = new TableCell();

                    TableCellProperties tableCellProperties5 = new TableCellProperties();
                    TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "4488", Type = TableWidthUnitValues.Dxa };

                    TableCellBorders tableCellBorders5 = new TableCellBorders();
                    LeftBorder leftBorder5 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };
                    BottomBorder bottomBorder5 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };

                    tableCellBorders5.Append(leftBorder5);
                    tableCellBorders5.Append(bottomBorder5);
                    Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

                    tableCellProperties5.Append(tableCellWidth5);
                    tableCellProperties5.Append(tableCellBorders5);
                    tableCellProperties5.Append(shading5);

                    Paragraph paragraph7 = new Paragraph();

                    ParagraphProperties paragraphProperties7 = new ParagraphProperties();
                    ParagraphStyleId paragraphStyleId7 = new ParagraphStyleId() { Val = "TableContents" };
                    BiDi biDi7 = new BiDi() { Val = false };
                    Justification justification7 = new Justification() { Val = JustificationValues.Left };

                    ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
                    RunFonts runFonts5 = new RunFonts() { Ascii = "Liberation Serif", HighAnsi = "Liberation Serif" };
                    Bold bold13 = new Bold() { Val = false };
                    Bold bold14 = new Bold() { Val = false };
                    BoldComplexScript boldComplexScript9 = new BoldComplexScript() { Val = false };
                    Italic italic13 = new Italic() { Val = false };
                    Italic italic14 = new Italic() { Val = false };
                    ItalicComplexScript italicComplexScript9 = new ItalicComplexScript() { Val = false };
                    Strike strike9 = new Strike() { Val = false };
                    DoubleStrike doubleStrike9 = new DoubleStrike() { Val = false };
                    Outline outline9 = new Outline() { Val = false };
                    Shadow shadow9 = new Shadow() { Val = false };
                    Color color9 = new Color() { Val = "000000" };
                    FontSize fontSize13 = new FontSize() { Val = "24" };
                    FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "24" };
                    Underline underline9 = new Underline() { Val = UnderlineValues.None };

                    paragraphMarkRunProperties7.Append(runFonts5);
                    paragraphMarkRunProperties7.Append(bold13);
                    paragraphMarkRunProperties7.Append(bold14);
                    paragraphMarkRunProperties7.Append(boldComplexScript9);
                    paragraphMarkRunProperties7.Append(italic13);
                    paragraphMarkRunProperties7.Append(italic14);
                    paragraphMarkRunProperties7.Append(italicComplexScript9);
                    paragraphMarkRunProperties7.Append(strike9);
                    paragraphMarkRunProperties7.Append(doubleStrike9);
                    paragraphMarkRunProperties7.Append(outline9);
                    paragraphMarkRunProperties7.Append(shadow9);
                    paragraphMarkRunProperties7.Append(color9);
                    paragraphMarkRunProperties7.Append(fontSize13);
                    paragraphMarkRunProperties7.Append(fontSizeComplexScript13);
                    paragraphMarkRunProperties7.Append(underline9);

                    paragraphProperties7.Append(paragraphStyleId7);
                    paragraphProperties7.Append(biDi7);
                    paragraphProperties7.Append(justification7);
                    paragraphProperties7.Append(paragraphMarkRunProperties7);

                    Run run7 = new Run();

                    RunProperties runProperties7 = new RunProperties();
                    Bold bold15 = new Bold() { Val = false };
                    BoldComplexScript boldComplexScript10 = new BoldComplexScript() { Val = false };
                    Italic italic15 = new Italic() { Val = false };
                    ItalicComplexScript italicComplexScript10 = new ItalicComplexScript() { Val = false };
                    Strike strike10 = new Strike() { Val = false };
                    DoubleStrike doubleStrike10 = new DoubleStrike() { Val = false };
                    Outline outline10 = new Outline() { Val = false };
                    Shadow shadow10 = new Shadow() { Val = false };
                    Color color10 = new Color() { Val = "000000" };
                    FontSize fontSize14 = new FontSize() { Val = "24" };
                    FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "24" };
                    Underline underline10 = new Underline() { Val = UnderlineValues.None };

                    runProperties7.Append(bold15);
                    runProperties7.Append(boldComplexScript10);
                    runProperties7.Append(italic15);
                    runProperties7.Append(italicComplexScript10);
                    runProperties7.Append(strike10);
                    runProperties7.Append(doubleStrike10);
                    runProperties7.Append(outline10);
                    runProperties7.Append(shadow10);
                    runProperties7.Append(color10);
                    runProperties7.Append(fontSize14);
                    runProperties7.Append(fontSizeComplexScript14);
                    runProperties7.Append(underline10);
                    Text text6 = new Text {Text = $"{order.Part.LegoElementId} {order.Part.LegoElementDescription}"};

                    run7.Append(runProperties7);
                    run7.Append(text6);

                    paragraph7.Append(paragraphProperties7);
                    paragraph7.Append(run7);

                    Paragraph paragraph8 = new Paragraph();

                    ParagraphProperties paragraphProperties8 = new ParagraphProperties();
                    ParagraphStyleId paragraphStyleId8 = new ParagraphStyleId() { Val = "TableContents" };
                    BiDi biDi8 = new BiDi() { Val = false };
                    Justification justification8 = new Justification() { Val = JustificationValues.Left };

                    ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
                    RunFonts runFonts6 = new RunFonts() { Ascii = "Liberation Serif", HighAnsi = "Liberation Serif" };
                    Bold bold16 = new Bold() { Val = false };
                    Bold bold17 = new Bold() { Val = false };
                    BoldComplexScript boldComplexScript11 = new BoldComplexScript() { Val = false };
                    Italic italic16 = new Italic() { Val = false };
                    Italic italic17 = new Italic() { Val = false };
                    ItalicComplexScript italicComplexScript11 = new ItalicComplexScript() { Val = false };
                    Strike strike11 = new Strike() { Val = false };
                    DoubleStrike doubleStrike11 = new DoubleStrike() { Val = false };
                    Outline outline11 = new Outline() { Val = false };
                    Shadow shadow11 = new Shadow() { Val = false };
                    Color color11 = new Color() { Val = "000000" };
                    FontSize fontSize15 = new FontSize() { Val = "24" };
                    FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "24" };
                    Underline underline11 = new Underline() { Val = UnderlineValues.None };

                    paragraphMarkRunProperties8.Append(runFonts6);
                    paragraphMarkRunProperties8.Append(bold16);
                    paragraphMarkRunProperties8.Append(bold17);
                    paragraphMarkRunProperties8.Append(boldComplexScript11);
                    paragraphMarkRunProperties8.Append(italic16);
                    paragraphMarkRunProperties8.Append(italic17);
                    paragraphMarkRunProperties8.Append(italicComplexScript11);
                    paragraphMarkRunProperties8.Append(strike11);
                    paragraphMarkRunProperties8.Append(doubleStrike11);
                    paragraphMarkRunProperties8.Append(outline11);
                    paragraphMarkRunProperties8.Append(shadow11);
                    paragraphMarkRunProperties8.Append(color11);
                    paragraphMarkRunProperties8.Append(fontSize15);
                    paragraphMarkRunProperties8.Append(fontSizeComplexScript15);
                    paragraphMarkRunProperties8.Append(underline11);

                    paragraphProperties8.Append(paragraphStyleId8);
                    paragraphProperties8.Append(biDi8);
                    paragraphProperties8.Append(justification8);
                    paragraphProperties8.Append(paragraphMarkRunProperties8);

                    Run run8 = new Run();

                    RunProperties runProperties8 = new RunProperties();
                    Bold bold18 = new Bold() { Val = false };
                    BoldComplexScript boldComplexScript12 = new BoldComplexScript() { Val = false };
                    Italic italic18 = new Italic() { Val = false };
                    ItalicComplexScript italicComplexScript12 = new ItalicComplexScript() { Val = false };
                    Strike strike12 = new Strike() { Val = false };
                    DoubleStrike doubleStrike12 = new DoubleStrike() { Val = false };
                    Outline outline12 = new Outline() { Val = false };
                    Shadow shadow12 = new Shadow() { Val = false };
                    Color color12 = new Color() { Val = "000000" };
                    FontSize fontSize16 = new FontSize() { Val = "24" };
                    FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "24" };
                    Underline underline12 = new Underline() { Val = UnderlineValues.None };

                    runProperties8.Append(bold18);
                    runProperties8.Append(boldComplexScript12);
                    runProperties8.Append(italic18);
                    runProperties8.Append(italicComplexScript12);
                    runProperties8.Append(strike12);
                    runProperties8.Append(doubleStrike12);
                    runProperties8.Append(outline12);
                    runProperties8.Append(shadow12);
                    runProperties8.Append(color12);
                    runProperties8.Append(fontSize16);
                    runProperties8.Append(fontSizeComplexScript16);
                    runProperties8.Append(underline12);
                    Text text7 = new Text {Text = $"{order.Part.LegoColorDescription} / {order.Part.BricklinkColorDescription}"};

                    run8.Append(runProperties8);
                    run8.Append(text7);

                    paragraph8.Append(paragraphProperties8);
                    paragraph8.Append(run8);

                    tableCell5.Append(tableCellProperties5);
                    tableCell5.Append(paragraph7);
                    tableCell5.Append(paragraph8);

                    TableCell tableCell6 = new TableCell();

                    TableCellProperties tableCellProperties6 = new TableCellProperties();
                    TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "3325", Type = TableWidthUnitValues.Dxa };

                    TableCellBorders tableCellBorders6 = new TableCellBorders();
                    LeftBorder leftBorder6 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };
                    BottomBorder bottomBorder6 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };
                    RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = 4U, Space = 0U };

                    tableCellBorders6.Append(leftBorder6);
                    tableCellBorders6.Append(bottomBorder6);
                    tableCellBorders6.Append(rightBorder2);
                    Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Fill = "auto" };

                    tableCellProperties6.Append(tableCellWidth6);
                    tableCellProperties6.Append(tableCellBorders6);
                    tableCellProperties6.Append(shading6);

                    Paragraph paragraph9 = new Paragraph();

                    ParagraphProperties paragraphProperties9 = new ParagraphProperties();
                    ParagraphStyleId paragraphStyleId9 = new ParagraphStyleId() { Val = "TableContents" };
                    BiDi biDi9 = new BiDi() { Val = false };
                    Justification justification9 = new Justification() { Val = JustificationValues.Left };

                    ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
                    RunFonts runFonts7 = new RunFonts() { Ascii = "Liberation Serif", HighAnsi = "Liberation Serif" };
                    Bold bold19 = new Bold() { Val = false };
                    Bold bold20 = new Bold() { Val = false };
                    BoldComplexScript boldComplexScript13 = new BoldComplexScript() { Val = false };
                    Italic italic19 = new Italic() { Val = false };
                    Italic italic20 = new Italic() { Val = false };
                    ItalicComplexScript italicComplexScript13 = new ItalicComplexScript() { Val = false };
                    Strike strike13 = new Strike() { Val = false };
                    DoubleStrike doubleStrike13 = new DoubleStrike() { Val = false };
                    Outline outline13 = new Outline() { Val = false };
                    Shadow shadow13 = new Shadow() { Val = false };
                    Color color13 = new Color() { Val = "000000" };
                    FontSize fontSize17 = new FontSize() { Val = "24" };
                    FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "24" };
                    Underline underline13 = new Underline() { Val = UnderlineValues.None };

                    paragraphMarkRunProperties9.Append(runFonts7);
                    paragraphMarkRunProperties9.Append(bold19);
                    paragraphMarkRunProperties9.Append(bold20);
                    paragraphMarkRunProperties9.Append(boldComplexScript13);
                    paragraphMarkRunProperties9.Append(italic19);
                    paragraphMarkRunProperties9.Append(italic20);
                    paragraphMarkRunProperties9.Append(italicComplexScript13);
                    paragraphMarkRunProperties9.Append(strike13);
                    paragraphMarkRunProperties9.Append(doubleStrike13);
                    paragraphMarkRunProperties9.Append(outline13);
                    paragraphMarkRunProperties9.Append(shadow13);
                    paragraphMarkRunProperties9.Append(color13);
                    paragraphMarkRunProperties9.Append(fontSize17);
                    paragraphMarkRunProperties9.Append(fontSizeComplexScript17);
                    paragraphMarkRunProperties9.Append(underline13);

                    paragraphProperties9.Append(paragraphStyleId9);
                    paragraphProperties9.Append(biDi9);
                    paragraphProperties9.Append(justification9);
                    paragraphProperties9.Append(paragraphMarkRunProperties9);

                    Run run9 = new Run();

                    RunProperties runProperties9 = new RunProperties();
                    Bold bold21 = new Bold() { Val = false };
                    BoldComplexScript boldComplexScript14 = new BoldComplexScript() { Val = false };
                    Italic italic21 = new Italic() { Val = false };
                    ItalicComplexScript italicComplexScript14 = new ItalicComplexScript() { Val = false };
                    Strike strike14 = new Strike() { Val = false };
                    DoubleStrike doubleStrike14 = new DoubleStrike() { Val = false };
                    Outline outline14 = new Outline() { Val = false };
                    Shadow shadow14 = new Shadow() { Val = false };
                    Color color14 = new Color() { Val = "000000" };
                    FontSize fontSize18 = new FontSize() { Val = "24" };
                    FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "24" };
                    Underline underline14 = new Underline() { Val = UnderlineValues.None };

                    runProperties9.Append(bold21);
                    runProperties9.Append(boldComplexScript14);
                    runProperties9.Append(italic21);
                    runProperties9.Append(italicComplexScript14);
                    runProperties9.Append(strike14);
                    runProperties9.Append(doubleStrike14);
                    runProperties9.Append(outline14);
                    runProperties9.Append(shadow14);
                    runProperties9.Append(color14);
                    runProperties9.Append(fontSize18);
                    runProperties9.Append(fontSizeComplexScript18);
                    runProperties9.Append(underline14);
                    Text text8 = new Text {Text = $"{order.Quantity}"};

                    run9.Append(runProperties9);
                    run9.Append(text8);

                    paragraph9.Append(paragraphProperties9);
                    paragraph9.Append(run9);

                    tableCell6.Append(tableCellProperties6);
                    tableCell6.Append(paragraph9);

                    tableRow2.Append(tableRowProperties2);
                    tableRow2.Append(tableCell4);
                    tableRow2.Append(tableCell5);
                    tableRow2.Append(tableCell6);

                    table1.Append(tableRow2);
                }

                body.Append(table1);

                Paragraph paragraph10 = new Paragraph();

                ParagraphProperties paragraphProperties10 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId10 = new ParagraphStyleId() { Val = "Normal" };
                BiDi biDi10 = new BiDi() { Val = false };
                Justification justification10 = new Justification() { Val = JustificationValues.Left };

                ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
                FontSize fontSize19 = new FontSize { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "32" };

                paragraphMarkRunProperties10.Append(fontSize19);
                paragraphMarkRunProperties10.Append(fontSizeComplexScript19);

                paragraphProperties10.Append(paragraphStyleId10);
                paragraphProperties10.Append(biDi10);
                paragraphProperties10.Append(justification10);
                paragraphProperties10.Append(paragraphMarkRunProperties10);

                Run run10 = new Run();

                RunProperties runProperties10 = new RunProperties();
                FontSize fontSize20 = new FontSize { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "32" };

                runProperties10.Append(fontSize20);
                runProperties10.Append(fontSizeComplexScript20);

                run10.Append(runProperties10);

                Run run11 = new Run();
                Break break2 = new Break { Type = BreakValues.Page };

                run11.Append(break2);

                paragraph10.Append(paragraphProperties10);
                paragraph10.Append(run10);
                paragraph10.Append(run11);

                body.Append(paragraph10);
            }

            document.Append(body);

            mainDocumentPart.Document = document;
        }

        private void GenerateStyleDefinitionsPartContent(StyleDefinitionsPart styleDefinitionsPart)
        {
            Styles styles1 = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            styles1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            styles1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts15 = new RunFonts() { Ascii = "Liberation Serif", HighAnsi = "Liberation Serif", EastAsia = "NSimSun", ComplexScript = "Arial" };
            Kern kern1 = new Kern() { Val = 2U };
            FontSize fontSize41 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "24" };
            Languages languages1 = new Languages() { Val = "en-US", EastAsia = "zh-CN", Bidi = "hi-IN" };

            runPropertiesBaseStyle1.Append(runFonts15);
            runPropertiesBaseStyle1.Append(kern1);
            runPropertiesBaseStyle1.Append(fontSize41);
            runPropertiesBaseStyle1.Append(fontSizeComplexScript41);
            runPropertiesBaseStyle1.Append(languages1);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);

            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            ParagraphPropertiesBaseStyle paragraphPropertiesBaseStyle1 = new ParagraphPropertiesBaseStyle();
            WidowControl widowControl1 = new WidowControl();

            paragraphPropertiesBaseStyle1.Append(widowControl1);

            paragraphPropertiesDefault1.Append(paragraphPropertiesBaseStyle1);

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal" };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            WidowControl widowControl2 = new WidowControl();
            Kinsoku kinsoku1 = new Kinsoku() { Val = true };
            OverflowPunctuation overflowPunctuation1 = new OverflowPunctuation() { Val = true };
            AutoSpaceDE autoSpaceDe1 = new AutoSpaceDE() { Val = true };
            BiDi biDi21 = new BiDi() { Val = false };

            styleParagraphProperties1.Append(widowControl2);
            styleParagraphProperties1.Append(kinsoku1);
            styleParagraphProperties1.Append(overflowPunctuation1);
            styleParagraphProperties1.Append(autoSpaceDe1);
            styleParagraphProperties1.Append(biDi21);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts16 = new RunFonts() { Ascii = "Liberation Serif", HighAnsi = "Liberation Serif", EastAsia = "NSimSun", ComplexScript = "Arial" };
            Color color29 = new Color() { Val = "auto" };
            Kern kern2 = new Kern() { Val = 2U };
            FontSize fontSize42 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "24" };
            Languages languages2 = new Languages() { Val = "en-US", EastAsia = "zh-CN", Bidi = "hi-IN" };

            styleRunProperties1.Append(runFonts16);
            styleRunProperties1.Append(color29);
            styleRunProperties1.Append(kern2);
            styleRunProperties1.Append(fontSize42);
            styleRunProperties1.Append(fontSizeComplexScript42);
            styleRunProperties1.Append(languages2);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

            Style style2 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading" };
            StyleName styleName2 = new StyleName() { Val = "Heading" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "TextBody" };
            PrimaryStyle primaryStyle2 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext() { Val = true };
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "240", After = "120" };

            styleParagraphProperties2.Append(keepNext1);
            styleParagraphProperties2.Append(spacingBetweenLines1);

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            RunFonts runFonts17 = new RunFonts() { Ascii = "Liberation Sans", HighAnsi = "Liberation Sans", EastAsia = "Microsoft YaHei", ComplexScript = "Arial" };
            FontSize fontSize43 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "28" };

            styleRunProperties2.Append(runFonts17);
            styleRunProperties2.Append(fontSize43);
            styleRunProperties2.Append(fontSizeComplexScript43);

            style2.Append(styleName2);
            style2.Append(basedOn1);
            style2.Append(nextParagraphStyle1);
            style2.Append(primaryStyle2);
            style2.Append(styleParagraphProperties2);
            style2.Append(styleRunProperties2);

            Style style3 = new Style() { Type = StyleValues.Paragraph, StyleId = "TextBody" };
            StyleName styleName3 = new StyleName() { Val = "Body Text" };
            BasedOn basedOn2 = new BasedOn() { Val = "Normal" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Before = "0", After = "140", Line = "276", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties3.Append(spacingBetweenLines2);
            StyleRunProperties styleRunProperties3 = new StyleRunProperties();

            style3.Append(styleName3);
            style3.Append(basedOn2);
            style3.Append(styleParagraphProperties3);
            style3.Append(styleRunProperties3);

            Style style4 = new Style() { Type = StyleValues.Paragraph, StyleId = "List" };
            StyleName styleName4 = new StyleName() { Val = "List" };
            BasedOn basedOn3 = new BasedOn() { Val = "TextBody" };
            StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            RunFonts runFonts18 = new RunFonts() { ComplexScript = "Arial" };

            styleRunProperties4.Append(runFonts18);

            style4.Append(styleName4);
            style4.Append(basedOn3);
            style4.Append(styleParagraphProperties4);
            style4.Append(styleRunProperties4);

            Style style5 = new Style() { Type = StyleValues.Paragraph, StyleId = "Caption" };
            StyleName styleName5 = new StyleName() { Val = "Caption" };
            BasedOn basedOn4 = new BasedOn() { Val = "Normal" };
            PrimaryStyle primaryStyle3 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties5 = new StyleParagraphProperties();
            SuppressLineNumbers suppressLineNumbers1 = new SuppressLineNumbers();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Before = "120", After = "120" };

            styleParagraphProperties5.Append(suppressLineNumbers1);
            styleParagraphProperties5.Append(spacingBetweenLines3);

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            RunFonts runFonts19 = new RunFonts() { ComplexScript = "Arial" };
            Italic italic43 = new Italic();
            ItalicComplexScript italicComplexScript29 = new ItalicComplexScript();
            FontSize fontSize44 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties5.Append(runFonts19);
            styleRunProperties5.Append(italic43);
            styleRunProperties5.Append(italicComplexScript29);
            styleRunProperties5.Append(fontSize44);
            styleRunProperties5.Append(fontSizeComplexScript44);

            style5.Append(styleName5);
            style5.Append(basedOn4);
            style5.Append(primaryStyle3);
            style5.Append(styleParagraphProperties5);
            style5.Append(styleRunProperties5);

            Style style6 = new Style() { Type = StyleValues.Paragraph, StyleId = "Index" };
            StyleName styleName6 = new StyleName() { Val = "Index" };
            BasedOn basedOn5 = new BasedOn() { Val = "Normal" };
            PrimaryStyle primaryStyle4 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties6 = new StyleParagraphProperties();
            SuppressLineNumbers suppressLineNumbers2 = new SuppressLineNumbers();

            styleParagraphProperties6.Append(suppressLineNumbers2);

            StyleRunProperties styleRunProperties6 = new StyleRunProperties();
            RunFonts runFonts20 = new RunFonts() { ComplexScript = "Arial" };

            styleRunProperties6.Append(runFonts20);

            style6.Append(styleName6);
            style6.Append(basedOn5);
            style6.Append(primaryStyle4);
            style6.Append(styleParagraphProperties6);
            style6.Append(styleRunProperties6);

            Style style7 = new Style() { Type = StyleValues.Paragraph, StyleId = "TableContents" };
            StyleName styleName7 = new StyleName() { Val = "Table Contents" };
            BasedOn basedOn6 = new BasedOn() { Val = "Normal" };
            PrimaryStyle primaryStyle5 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties7 = new StyleParagraphProperties();
            SuppressLineNumbers suppressLineNumbers3 = new SuppressLineNumbers();

            styleParagraphProperties7.Append(suppressLineNumbers3);
            StyleRunProperties styleRunProperties7 = new StyleRunProperties();

            style7.Append(styleName7);
            style7.Append(basedOn6);
            style7.Append(primaryStyle5);
            style7.Append(styleParagraphProperties7);
            style7.Append(styleRunProperties7);

            styles1.Append(docDefaults1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);
            styles1.Append(style6);
            styles1.Append(style7);

            styleDefinitionsPart.Styles = styles1;

        }

        private void GenerateImagePartContent(ImagePart imagePart, string elementId)
        {
            Stream image = File.Exists($"./Images/{elementId}.png") ?
                new MemoryStream(File.ReadAllBytes($"./Images/{elementId}.png")) :
                new MemoryStream(File.ReadAllBytes($"./Images/ImageNotAvailable.png"));

            imagePart.FeedData(image);
            image.Close();
        }

        private void GenerateFontTablePartContent(FontTablePart fontTablePart)
        {
            Fonts fonts1 = new Fonts();
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            Font font1 = new Font() { Name = "Times New Roman" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };

            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);

            Font font2 = new Font() { Name = "Symbol" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "02" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };

            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);

            Font font3 = new Font() { Name = "Arial" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Variable };

            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);

            Font font4 = new Font() { Name = "Liberation Serif" };
            AltName altName1 = new AltName() { Val = "Times New Roman" };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch4 = new Pitch() { Val = FontPitchValues.Variable };

            font4.Append(altName1);
            font4.Append(fontCharSet4);
            font4.Append(fontFamily4);
            font4.Append(pitch4);

            Font font5 = new Font() { Name = "Liberation Sans" };
            AltName altName2 = new AltName() { Val = "Arial" };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily5 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch5 = new Pitch() { Val = FontPitchValues.Variable };

            font5.Append(altName2);
            font5.Append(fontCharSet5);
            font5.Append(fontFamily5);
            font5.Append(pitch5);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);

            fontTablePart.Fonts = fonts1;
        }

        private void GenerateDocumentSettingsPartContent(DocumentSettingsPart documentSettingsPart)
        {
            Settings settings = new Settings();
            settings.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            Zoom zoom = new Zoom{ Percent = "100" };
            DefaultTabStop defaultTabStop = new DefaultTabStop { Val = 709 };

            settings.Append(zoom);
            settings.Append(defaultTabStop);

            documentSettingsPart.Settings = settings;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "LUGTools/Picklists";
            document.PackageProperties.Title = "KLUG Per Person Picklist";
            document.PackageProperties.Subject = "LUGBULK 2020";
            document.PackageProperties.Description = "KLUG Per Person Picklist for LUGBULK 2020";
            document.PackageProperties.Revision = "1";
            document.PackageProperties.Created = System.DateTime.Now;
            document.PackageProperties.Modified = System.DateTime.Now;
            document.PackageProperties.LastModifiedBy = "LUGTools/Picklists";
            document.PackageProperties.Language = "en-US";
        }

        private readonly PicklistGeneratorOptions options;
    }
}