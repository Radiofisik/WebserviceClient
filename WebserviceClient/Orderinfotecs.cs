using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using M = DocumentFormat.OpenXml.Math;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using System;

namespace GeneratedCode
{
    public class GeneratedClass
    {
        /// <summary>
        /// ///////////////////////////
        /// </summary>
        public WebserviceClient.srv.order orderi;/////////////////////////////////////////////





        // Creates a WordprocessingDocument.
        public void CreatePackage(string filePath)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(WordprocessingDocument document)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            MainDocumentPart mainDocumentPart1 = document.AddMainDocumentPart();
            GenerateMainDocumentPart1Content(mainDocumentPart1);

            StylesWithEffectsPart stylesWithEffectsPart1 = mainDocumentPart1.AddNewPart<StylesWithEffectsPart>("rId3");
            GenerateStylesWithEffectsPart1Content(stylesWithEffectsPart1);

            EndnotesPart endnotesPart1 = mainDocumentPart1.AddNewPart<EndnotesPart>("rId7");
            GenerateEndnotesPart1Content(endnotesPart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId2");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            NumberingDefinitionsPart numberingDefinitionsPart1 = mainDocumentPart1.AddNewPart<NumberingDefinitionsPart>("rId1");
            GenerateNumberingDefinitionsPart1Content(numberingDefinitionsPart1);

            FootnotesPart footnotesPart1 = mainDocumentPart1.AddNewPart<FootnotesPart>("rId6");
            GenerateFootnotesPart1Content(footnotesPart1);

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId5");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId10");
            GenerateThemePart1Content(themePart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId4");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId9");
            GenerateFontTablePart1Content(fontTablePart1);

            mainDocumentPart1.AddHyperlinkRelationship(new System.Uri("mailto:Soft@infotecs.ru", System.UriKind.Absolute), true, "rId8");
            SetPackageProperties(document);
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Template template1 = new Ap.Template();
            template1.Text = "Normal.dotm";
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "0";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "1";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "265";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "1517";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office Word";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.Lines lines1 = new Ap.Lines();
            lines1.Text = "12";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "3";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Название";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "1";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)1U };
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "";

            vTVector2.Append(vTLPSTR2);

            titlesOfParts1.Append(vTVector2);
            Ap.Company company1 = new Ap.Company();
            company1.Text = " \"КГ НИЦ\"";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "1779";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";

            Ap.HyperlinkList hyperlinkList1 = new Ap.HyperlinkList();

            Vt.VTVector vTVector3 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)6U };

            Vt.Variant variant3 = new Vt.Variant();
            Vt.VTInt32 vTInt322 = new Vt.VTInt32();
            vTInt322.Text = "5374055";

            variant3.Append(vTInt322);

            Vt.Variant variant4 = new Vt.Variant();
            Vt.VTInt32 vTInt323 = new Vt.VTInt32();
            vTInt323.Text = "0";

            variant4.Append(vTInt323);

            Vt.Variant variant5 = new Vt.Variant();
            Vt.VTInt32 vTInt324 = new Vt.VTInt32();
            vTInt324.Text = "0";

            variant5.Append(vTInt324);

            Vt.Variant variant6 = new Vt.Variant();
            Vt.VTInt32 vTInt325 = new Vt.VTInt32();
            vTInt325.Text = "5";

            variant6.Append(vTInt325);

            Vt.Variant variant7 = new Vt.Variant();
            Vt.VTLPWSTR vTLPWSTR1 = new Vt.VTLPWSTR();
            vTLPWSTR1.Text = "mailto:Soft@infotecs.ru";

            variant7.Append(vTLPWSTR1);

            Vt.Variant variant8 = new Vt.Variant();
            Vt.VTLPWSTR vTLPWSTR2 = new Vt.VTLPWSTR();
            vTLPWSTR2.Text = "";

            variant8.Append(vTLPWSTR2);

            vTVector3.Append(variant3);
            vTVector3.Append(variant4);
            vTVector3.Append(variant5);
            vTVector3.Append(variant6);
            vTVector3.Append(variant7);
            vTVector3.Append(variant8);

            hyperlinkList1.Append(vTVector3);
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "14.0000";

            properties1.Append(template1);
            properties1.Append(totalTime1);
            properties1.Append(pages1);
            properties1.Append(words1);
            properties1.Append(characters1);
            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(lines1);
            properties1.Append(paragraphs1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(charactersWithSpaces1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinkList1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        // Generates content of mainDocumentPart1.
        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1)
        {
            Document document1 = new Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            document1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            document1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            document1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            document1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            document1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            document1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Body body1 = new Body();

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00831455", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00831455" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "1" };

            paragraphProperties1.Append(paragraphStyleId1);

            Run run1 = new Run() { RsidRunProperties = "001555D6" };
            Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text1.Text = "ЗАЯВКА на ";

            run1.Append(text1);

            Run run2 = new Run() { RsidRunAddition = "001555D6" };
            Text text2 = new Text();
            text2.Text = "приобретение";

            run2.Append(text2);

            Run run3 = new Run() { RsidRunAddition = "00F444F2" };
            Text text3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text3.Text = " ПО";

            run3.Append(text3);

            Run run4 = new Run() { RsidRunProperties = "001555D6" };
            Text text4 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text4.Text = " ";

            run4.Append(text4);

            Run run5 = new Run() { RsidRunProperties = "001555D6" };

            RunProperties runProperties1 = new RunProperties();
            Languages languages1 = new Languages() { Val = "en-US" };

            runProperties1.Append(languages1);
            Text text5 = new Text();
            text5.Text = "ViPNet";

            run5.Append(runProperties1);
            run5.Append(text5);

            Run run6 = new Run() { RsidRunProperties = "001555D6", RsidRunAddition = "001D5FD5" };
            Text text6 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text6.Text = " ";

            run6.Append(text6);

            Run run7 = new Run() { RsidRunAddition = "001555D6" };
            Text text7 = new Text();
            text7.Text = "Партнерами";

            run7.Append(text7);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            paragraph1.Append(run2);
            paragraph1.Append(run3);
            paragraph1.Append(run4);
            paragraph1.Append(run5);
            paragraph1.Append(run6);
            paragraph1.Append(run7);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00F444F2", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00F444F2" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties2.Append(paragraphStyleId2);

            paragraph2.Append(paragraphProperties2);

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableWidth tableWidth1 = new TableWidth() { Width = "10491", Type = TableWidthUnitValues.Dxa };
            TableIndentation tableIndentation1 = new TableIndentation() { Width = -885, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);
            TableLook tableLook1 = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableIndentation1);
            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "10491" };

            tableGrid1.Append(gridColumn1);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "000E0BE7", RsidTableRowAddition = "00F444F2", RsidTableRowProperties = "000E0BE7" };

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "10491", Type = TableWidthUnitValues.Dxa };

            tableCellProperties1.Append(tableCellWidth1);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "000E0BE7", RsidParagraphAddition = "00F444F2", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00F444F2" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties3.Append(paragraphStyleId3);

            Run run8 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text8 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text8.Text = "Наименование организации Партнера на кого необходимо выставить счет в рамках данного ";

            run8.Append(text8);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run8);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "000E0BE7", RsidParagraphAddition = "00F444F2", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00F444F2" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties4.Append(paragraphStyleId4);

            Run run9 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text9 = new Text();
            text9.Text = "запроса (";

            run9.Append(text9);

            Run run10 = new Run() { RsidRunAddition = "00F15D05" };
            Text text10 = new Text();
            text10.Text = "е";

            run10.Append(text10);

            Run run11 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text11 = new Text();
            text11.Text = "сли филиал";

            run11.Append(text11);

            Run run12 = new Run() { RsidRunProperties = "000E0BE7", RsidRunAddition = "006E1D99" };
            Text text12 = new Text();
            text12.Text = ",";

            run12.Append(text12);

            Run run13 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text13 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text13.Text = " то указать какой именно): ";

            run13.Append(text13);

            Run run14 = new Run() { RsidRunAddition = "00CC6D34" };
            Text text14 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text14.Text = "ООО «Агентство информационной безопасности» Юридический и почтовый адрес: 236022, г. ";

            run14.Append(text14);

            Run run15 = new Run() { RsidRunProperties = "00D37E6B", RsidRunAddition = "00CC6D34" };
            Text text15 = new Text();
            text15.Text = "Калининград";

            run15.Append(text15);
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };

            Run run16 = new Run() { RsidRunAddition = "00CC6D34" };
            Text text16 = new Text();
            text16.Text = ", ул. Генделя, 5, этаж 5, офисы 31 - 45";

            run16.Append(text16);
           

            Run run18 = new Run() { RsidRunAddition = "00CC6D34" };
            Text text18 = new Text();
            text18.Text = "Тел./факс (4012) 99-22-86, 99-22-63";

            run18.Append(text18);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run9);
            paragraph4.Append(run10);
            paragraph4.Append(run11);
            paragraph4.Append(run12);
            paragraph4.Append(run13);
            paragraph4.Append(run14);
            paragraph4.Append(run15);
            paragraph4.Append(bookmarkStart1);
            paragraph4.Append(bookmarkEnd1);
            paragraph4.Append(run16);
            paragraph4.Append(run18);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "00CC6D34", RsidParagraphAddition = "00F444F2", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00F444F2" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId5 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties5.Append(paragraphStyleId5);

            Run run19 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text19 = new Text();
            text19.Text = "Дата заполнения заявки:";

            run19.Append(text19);

            Run run20 = new Run() { RsidRunAddition = "00A72F6F" };
            Text text20 = new Text();
            text20.Text = DateTime.Today.ToShortDateString();

            run20.Append(text20);


            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run19);
            paragraph5.Append(run20);
        

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph3);
            tableCell1.Append(paragraph4);
            tableCell1.Append(paragraph5);

            tableRow1.Append(tableCell1);

            TableRow tableRow2 = new TableRow() { RsidTableRowMarkRevision = "000E0BE7", RsidTableRowAddition = "00F444F2", RsidTableRowProperties = "000E0BE7" };

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "10491", Type = TableWidthUnitValues.Dxa };

            tableCellProperties2.Append(tableCellWidth2);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "00CC6D34", RsidParagraphAddition = "00F444F2", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00F444F2" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId6 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties6.Append(paragraphStyleId6);

            Run run24 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text24 = new Text();
            text24.Text = "ФИО,  контактная информация лица заполняющего заявку:";

            run24.Append(text24);

            Run run25 = new Run() { RsidRunProperties = "00CC6D34", RsidRunAddition = "00CC6D34" };
            Text text25 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text25.Text = " ";

            run25.Append(text25);

            Run run26 = new Run() { RsidRunAddition = "00D37E6B" };
            Text text26 = new Text();
            text26.Text = "Егоркин";

            run26.Append(text26);

            Run run27 = new Run() { RsidRunProperties = "00CC6D34", RsidRunAddition = "00CC6D34" };
            Text text27 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text27.Text = " Александр ";

            run27.Append(text27);

            Run run28 = new Run() { RsidRunAddition = "00D37E6B" };
            Text text28 = new Text();
            text28.Text = "Евгеньевич";

            run28.Append(text28);

            Run run29 = new Run() { RsidRunAddition = "00CC6D34" };
            Text text29 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text29.Text = " 8-4012-99-22-65 kgnic@mail.ru";

            run29.Append(text29);

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(run24);
            paragraph6.Append(run25);
            paragraph6.Append(run26);
            paragraph6.Append(run27);
            paragraph6.Append(run28);
            paragraph6.Append(run29);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph6);

            tableRow2.Append(tableCell2);

            TableRow tableRow3 = new TableRow() { RsidTableRowMarkRevision = "000E0BE7", RsidTableRowAddition = "00F444F2", RsidTableRowProperties = "000E0BE7" };

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "10491", Type = TableWidthUnitValues.Dxa };

            tableCellProperties3.Append(tableCellWidth3);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "000E0BE7", RsidParagraphAddition = "00F444F2", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00F444F2" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId7 = new ParagraphStyleId() { Val = "a3" };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            Languages languages4 = new Languages() { Val = "ru" };

            paragraphMarkRunProperties1.Append(languages4);

            paragraphProperties7.Append(paragraphStyleId7);
            paragraphProperties7.Append(paragraphMarkRunProperties1);
            ProofError proofError3 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run30 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text30 = new Text();
            text30.Text = "Дополнительная информация";

            run30.Append(text30);

            Run run31 = new Run() { RsidRunAddition = "00F15D05" };
            Text text31 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text31.Text = " ";

            run31.Append(text31);

            Run run32 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text32 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text32.Text = " (";

            run32.Append(text32);

            Run run33 = new Run() { RsidRunAddition = "00F15D05" };
            Text text33 = new Text();
            text33.Text = "н";

            run33.Append(text33);

            Run run34 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text34 = new Text();
            text34.Text = "омер договора, предполагаемая скидка, форма оплаты, не стандартный адрес доставки, и. т. д.)";

            run34.Append(text34);

            Run run35 = new Run() { RsidRunAddition = "00F15D05" };
            Text text35 = new Text();
            text35.Text = ":";

            run35.Append(text35);

            Run run36 = new Run() { RsidRunAddition = "00CC6D34" };
            Text text36 = new Text();
            text36.Text = "-";

            run36.Append(text36);
            ProofError proofError4 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            paragraph7.Append(paragraphProperties7);
            paragraph7.Append(proofError3);
            paragraph7.Append(run30);
            paragraph7.Append(run31);
            paragraph7.Append(run32);
            paragraph7.Append(run33);
            paragraph7.Append(run34);
            paragraph7.Append(run35);
            paragraph7.Append(run36);
            paragraph7.Append(proofError4);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph7);

            tableRow3.Append(tableCell3);

            TableRow tableRow4 = new TableRow() { RsidTableRowMarkRevision = "000E0BE7", RsidTableRowAddition = "007F46F0", RsidTableRowProperties = "000E0BE7" };

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "10491", Type = TableWidthUnitValues.Dxa };

            tableCellProperties4.Append(tableCellWidth4);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "000E0BE7", RsidParagraphAddition = "007F46F0", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "007F46F0" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId8 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties8.Append(paragraphStyleId8);

            Run run37 = new Run();
            Text text37 = new Text();
            text37.Text = "Требуется ли доставка курьерской службой (";

            run37.Append(text37);

            Run run38 = new Run() { RsidRunAddition = "00F15D05" };
            Text text38 = new Text();
            text38.Text = "п";

            run38.Append(text38);

            Run run39 = new Run();
            Text text39 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text39.Text = "о мелким заказам доставка будет включена в счет на оплату отдельной строкой, цена доставки соответствует ";

            run39.Append(text39);

            Run run40 = new Run() { RsidRunAddition = "00A72F6F" };
            Text text40 = new Text();
            text40.Text = "прайсу";

            run40.Append(text40);

            Run run41 = new Run();
            Text text41 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text41.Text = " курьерской службы Пони-Экспресс)";

            run41.Append(text41);

            Run run42 = new Run() { RsidRunAddition = "00F15D05" };
            Text text42 = new Text();
            text42.Text = ":";

            run42.Append(text42);

            Run run43 = new Run() { RsidRunAddition = "00CC6D34" };
            Text text43 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text43.Text = " ";

            run43.Append(text43);

            Run run44 = new Run() { RsidRunProperties = "00A72F6F", RsidRunAddition = "00A72F6F" };
            Text text44 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text44.Text = "не ";

            run44.Append(text44);

            Run run45 = new Run() { RsidRunAddition = "00CC6D34" };
            Text text45 = new Text();
            text45.Text = "требуется доставка";

            run45.Append(text45);

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(run37);
            paragraph8.Append(run38);
            paragraph8.Append(run39);
            paragraph8.Append(run40);
            paragraph8.Append(run41);
            paragraph8.Append(run42);
            paragraph8.Append(run43);
            paragraph8.Append(run44);
            paragraph8.Append(run45);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph8);

            tableRow4.Append(tableCell4);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);
            table1.Append(tableRow3);
            table1.Append(tableRow4);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "001555D6", RsidParagraphAddition = "00831455", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00831455" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId9 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties9.Append(paragraphStyleId9);

            paragraph9.Append(paragraphProperties9);








            Table table2 = new Table();

            TableProperties tableProperties2 = new TableProperties();
            TableWidth tableWidth2 = new TableWidth() { Width = "10331", Type = TableWidthUnitValues.Dxa };
            TableIndentation tableIndentation2 = new TableIndentation() { Width = -908, Type = TableWidthUnitValues.Dxa };
            TableLayout tableLayout1 = new TableLayout() { Type = TableLayoutValues.Fixed };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 10, Type = TableWidthValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 10, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);
            TableLook tableLook2 = new TableLook() { Val = "0000", FirstRow = false, LastRow = false, FirstColumn = false, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = false };

            tableProperties2.Append(tableWidth2);
            tableProperties2.Append(tableIndentation2);
            tableProperties2.Append(tableLayout1);
            tableProperties2.Append(tableCellMarginDefault1);
            tableProperties2.Append(tableLook2);

            TableGrid tableGrid2 = new TableGrid();
            GridColumn gridColumn2 = new GridColumn() { Width = "610" };
            GridColumn gridColumn3 = new GridColumn() { Width = "2151" };
            GridColumn gridColumn4 = new GridColumn() { Width = "1008" };
            GridColumn gridColumn5 = new GridColumn() { Width = "1260" };
            GridColumn gridColumn6 = new GridColumn() { Width = "2126" };
            GridColumn gridColumn7 = new GridColumn() { Width = "3176" };

            tableGrid2.Append(gridColumn2);
            tableGrid2.Append(gridColumn3);
            tableGrid2.Append(gridColumn4);
            tableGrid2.Append(gridColumn5);
            tableGrid2.Append(gridColumn6);
            tableGrid2.Append(gridColumn7);

            TableRow tableRow5 = new TableRow() { RsidTableRowMarkRevision = "001555D6", RsidTableRowAddition = "00831455", RsidTableRowProperties = "002D5ECD" };

            TablePropertyExceptions tablePropertyExceptions1 = new TablePropertyExceptions();

            TableCellMarginDefault tableCellMarginDefault2 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };

            tableCellMarginDefault2.Append(topMargin1);
            tableCellMarginDefault2.Append(bottomMargin1);

            tablePropertyExceptions1.Append(tableCellMarginDefault2);

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)1459U };

            tableRowProperties1.Append(tableRowHeight1);

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "610", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders1 = new TableCellBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders1.Append(topBorder2);
            tableCellBorders1.Append(leftBorder2);
            tableCellBorders1.Append(bottomBorder2);
            tableCellBorders1.Append(rightBorder2);
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            tableCellProperties5.Append(tableCellWidth5);
            tableCellProperties5.Append(tableCellBorders1);
            tableCellProperties5.Append(shading1);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "000E0BE7", RsidParagraphAddition = "00831455", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00831455" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId10 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties10.Append(paragraphStyleId10);

            Run run46 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text46 = new Text();
            text46.Text = "№";

            run46.Append(text46);

            paragraph10.Append(paragraphProperties10);
            paragraph10.Append(run46);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph10);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "2151", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders2 = new TableCellBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders2.Append(topBorder3);
            tableCellBorders2.Append(leftBorder3);
            tableCellBorders2.Append(bottomBorder3);
            tableCellBorders2.Append(rightBorder3);
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            tableCellProperties6.Append(tableCellWidth6);
            tableCellProperties6.Append(tableCellBorders2);
            tableCellProperties6.Append(shading2);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "001555D6", RsidParagraphAddition = "00831455", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00831455" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId11 = new ParagraphStyleId() { Val = "a3" };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            Languages languages5 = new Languages() { Val = "en-US" };

            paragraphMarkRunProperties2.Append(languages5);

            paragraphProperties11.Append(paragraphStyleId11);
            paragraphProperties11.Append(paragraphMarkRunProperties2);

            Run run47 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text47 = new Text();
            text47.Text = "Наименование продукта";

            run47.Append(text47);

            Run run48 = new Run() { RsidRunProperties = "000E0BE7", RsidRunAddition = "001555D6" };
            Text text48 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text48.Text = " ";

            run48.Append(text48);

            Run run49 = new Run() { RsidRunAddition = "001555D6" };

            RunProperties runProperties4 = new RunProperties();
            Languages languages6 = new Languages() { Val = "en-US" };

            runProperties4.Append(languages6);
            Text text49 = new Text();
            text49.Text = "ViPNet";

            run49.Append(runProperties4);
            run49.Append(text49);

            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(run47);
            paragraph11.Append(run48);
            paragraph11.Append(run49);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph11);

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "1008", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders3 = new TableCellBorders();
            TopBorder topBorder4 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder4 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders3.Append(topBorder4);
            tableCellBorders3.Append(leftBorder4);
            tableCellBorders3.Append(bottomBorder4);
            tableCellBorders3.Append(rightBorder4);
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            tableCellProperties7.Append(tableCellWidth7);
            tableCellProperties7.Append(tableCellBorders3);
            tableCellProperties7.Append(shading3);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphMarkRevision = "000E0BE7", RsidParagraphAddition = "00831455", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00831455" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId12 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties12.Append(paragraphStyleId12);

            Run run50 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text50 = new Text();
            text50.Text = "Кол-во лицензий";

            run50.Append(text50);

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(run50);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph12);

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "1260", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders4 = new TableCellBorders();
            TopBorder topBorder5 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder5 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder5 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder5 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders4.Append(topBorder5);
            tableCellBorders4.Append(leftBorder5);
            tableCellBorders4.Append(bottomBorder5);
            tableCellBorders4.Append(rightBorder5);
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            tableCellProperties8.Append(tableCellWidth8);
            tableCellProperties8.Append(tableCellBorders4);
            tableCellProperties8.Append(shading4);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphMarkRevision = "001555D6", RsidParagraphAddition = "00831455", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "001555D6" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId13 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties13.Append(paragraphStyleId13);

            Run run51 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text51 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text51.Text = "Кол-во ";

            run51.Append(text51);

            Run run52 = new Run();

            RunProperties runProperties5 = new RunProperties();
            Languages languages7 = new Languages() { Val = "en-US" };

            runProperties5.Append(languages7);
            Text text52 = new Text();
            text52.Text = "CD";

            run52.Append(runProperties5);
            run52.Append(text52);

            Run run53 = new Run() { RsidRunProperties = "001555D6" };
            Text text53 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text53.Text = " ";

            run53.Append(text53);

            Run run54 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text54 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text54.Text = "дисков с дистрибутивами ";

            run54.Append(text54);
            ProofError proofError5 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run55 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text55 = new Text();
            text55.Text = "ПО";

            run55.Append(text55);
            ProofError proofError6 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run56 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text56 = new Text();
            text56.Text = ".";

            run56.Append(text56);

            paragraph13.Append(paragraphProperties13);
            paragraph13.Append(run51);
            paragraph13.Append(run52);
            paragraph13.Append(run53);
            paragraph13.Append(run54);
            paragraph13.Append(proofError5);
            paragraph13.Append(run55);
            paragraph13.Append(proofError6);
            paragraph13.Append(run56);

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph13);

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "2126", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders5 = new TableCellBorders();
            TopBorder topBorder6 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder6 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder6 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder6 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders5.Append(topBorder6);
            tableCellBorders5.Append(leftBorder6);
            tableCellBorders5.Append(bottomBorder6);
            tableCellBorders5.Append(rightBorder6);
            Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            tableCellProperties9.Append(tableCellWidth9);
            tableCellProperties9.Append(tableCellBorders5);
            tableCellProperties9.Append(shading5);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphMarkRevision = "006E1D99", RsidParagraphAddition = "00831455", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "006E1D99" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId14 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties14.Append(paragraphStyleId14);

            Run run57 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text57 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text57.Text = "Номер сети ";

            run57.Append(text57);

            Run run58 = new Run();

            RunProperties runProperties6 = new RunProperties();
            Languages languages8 = new Languages() { Val = "en-US" };

            runProperties6.Append(languages8);
            Text text58 = new Text();
            text58.Text = "ViPNet";

            run58.Append(runProperties6);
            run58.Append(text58);

            Run run59 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text59 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text59.Text = " версия";

            run59.Append(text59);

       

            Run run61 = new Run() { RsidRunProperties = "000E0BE7", RsidRunAddition = "00F444F2" };
            Text text61 = new Text();
            text61.Text = ",";

            run61.Append(text61);

            Run run62 = new Run() { RsidRunProperties = "000E0BE7", RsidRunAddition = "001555D6" };
            Text text62 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text62.Text = " ";

            run62.Append(text62);

           
            Run run64 = new Run() { RsidRunProperties = "000E0BE7", RsidRunAddition = "001555D6" };
            Text text64 = new Text();
            text64.Text = "требования к сертификации (КС 2, КС 3)";

            run64.Append(text64);

            Run run65 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text65 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text65.Text = " ";

            run65.Append(text65);

            paragraph14.Append(paragraphProperties14);
            paragraph14.Append(run57);
            paragraph14.Append(run58);
            paragraph14.Append(run59);
         
            paragraph14.Append(run61);
            paragraph14.Append(run62);
            paragraph14.Append(run64);
            paragraph14.Append(run65);

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(paragraph14);

            TableCell tableCell10 = new TableCell();

            TableCellProperties tableCellProperties10 = new TableCellProperties();
            TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "3176", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders6 = new TableCellBorders();
            TopBorder topBorder7 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder7 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder7 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder7 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableCellBorders6.Append(topBorder7);
            tableCellBorders6.Append(leftBorder7);
            tableCellBorders6.Append(bottomBorder7);
            tableCellBorders6.Append(rightBorder7);
            Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            tableCellProperties10.Append(tableCellWidth10);
            tableCellProperties10.Append(tableCellBorders6);
            tableCellProperties10.Append(shading6);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "000E0BE7", RsidParagraphAddition = "00831455", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00F444F2" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId15 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties15.Append(paragraphStyleId15);

            Run run66 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text66 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text66.Text = "Данные о конечном пользователе ";

            run66.Append(text66);
            ProofError proofError7 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run67 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text67 = new Text();
            text67.Text = "ПО";

            run67.Append(text67);
            ProofError proofError8 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run68 = new Run() { RsidRunProperties = "000E0BE7" };
            Text text68 = new Text();
            text68.Text = ": Название организации, почтовый адрес, ИНН/КПП.";

            run68.Append(text68);

            paragraph15.Append(paragraphProperties15);
            paragraph15.Append(run66);
            paragraph15.Append(proofError7);
            paragraph15.Append(run67);
            paragraph15.Append(proofError8);
            paragraph15.Append(run68);

            tableCell10.Append(tableCellProperties10);
            tableCell10.Append(paragraph15);

            tableRow5.Append(tablePropertyExceptions1);
            tableRow5.Append(tableRowProperties1);
            tableRow5.Append(tableCell5);
            tableRow5.Append(tableCell6);
            tableRow5.Append(tableCell7);
            tableRow5.Append(tableCell8);
            tableRow5.Append(tableCell9);
            tableRow5.Append(tableCell10);

   
            



            table2.Append(tableProperties2);
            table2.Append(tableGrid2);



            table2.Append(tableRow5);
   


            // Заполнение таблицы ////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////////////////////
            for (int i = 0; i < orderi.items.Length; i++)
            {


                TableRow tableRow6 = new TableRow() { RsidTableRowMarkRevision = "001555D6", RsidTableRowAddition = "00831455", RsidTableRowProperties = "006E1D99" };

                TablePropertyExceptions tablePropertyExceptions2 = new TablePropertyExceptions();

                TableCellMarginDefault tableCellMarginDefault3 = new TableCellMarginDefault();
                TopMargin topMargin2 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
                BottomMargin bottomMargin2 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };

                tableCellMarginDefault3.Append(topMargin2);
                tableCellMarginDefault3.Append(bottomMargin2);

                tablePropertyExceptions2.Append(tableCellMarginDefault3);

                TableRowProperties tableRowProperties2 = new TableRowProperties();
                TableRowHeight tableRowHeight2 = new TableRowHeight() { Val = (UInt32Value)490U };

                tableRowProperties2.Append(tableRowHeight2);

                TableCell tableCell11 = new TableCell();

                TableCellProperties tableCellProperties11 = new TableCellProperties();
                TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "610", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders7 = new TableCellBorders();
                TopBorder topBorder8 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder8 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder8 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder8 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

                tableCellBorders7.Append(topBorder8);
                tableCellBorders7.Append(leftBorder8);
                tableCellBorders7.Append(bottomBorder8);
                tableCellBorders7.Append(rightBorder8);
                Shading shading7 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

                tableCellProperties11.Append(tableCellWidth11);
                tableCellProperties11.Append(tableCellBorders7);
                tableCellProperties11.Append(shading7);

                Paragraph paragraph16 = new Paragraph() { RsidParagraphMarkRevision = "001555D6", RsidParagraphAddition = "00831455", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00831455" };

                ParagraphProperties paragraphProperties16 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId16 = new ParagraphStyleId() { Val = "a3" };

                paragraphProperties16.Append(paragraphStyleId16);

                /////////////////////////
                Run run100 = new Run() { RsidRunProperties = "001555D6" };

                RunProperties runProperties100 = new RunProperties();
                RunFonts runFonts100 = new RunFonts() { EastAsia = "Times New Roman" };
                Languages languages100 = new Languages() { EastAsia = "ru-RU" };

                runProperties100.Append(runFonts100);
                runProperties100.Append(languages100);
                Text text100 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text100.Text = (i+1).ToString();/////////

                run100.Append(runProperties100);
                run100.Append(text100);
                ///////////////////////

                paragraph16.Append(paragraphProperties16);
                paragraph16.Append(run100);

                tableCell11.Append(tableCellProperties11);
                tableCell11.Append(paragraph16);

                TableCell tableCell12 = new TableCell();

                TableCellProperties tableCellProperties12 = new TableCellProperties();
                TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "2151", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders8 = new TableCellBorders();
                TopBorder topBorder9 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder9 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder9 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder9 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

                tableCellBorders8.Append(topBorder9);
                tableCellBorders8.Append(leftBorder9);
                tableCellBorders8.Append(bottomBorder9);
                tableCellBorders8.Append(rightBorder9);
                Shading shading8 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

                tableCellProperties12.Append(tableCellWidth12);
                tableCellProperties12.Append(tableCellBorders8);
                tableCellProperties12.Append(shading8);

                Paragraph paragraph17 = new Paragraph() { RsidParagraphMarkRevision = "00CC6D34", RsidParagraphAddition = "00831455", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00E46DD7" };

                ParagraphProperties paragraphProperties17 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId17 = new ParagraphStyleId() { Val = "a3" };

                paragraphProperties17.Append(paragraphStyleId17);

                Run run69 = new Run() { RsidRunProperties = "001555D6" };

                RunProperties runProperties7 = new RunProperties();
                RunFonts runFonts1 = new RunFonts() { EastAsia = "Times New Roman" };
                Languages languages9 = new Languages() { EastAsia = "ru-RU" };

                runProperties7.Append(runFonts1);
                runProperties7.Append(languages9);
                Text text69 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text69.Text = orderi.items[i].Product;/////////

                run69.Append(runProperties7);
                run69.Append(text69);

                paragraph17.Append(paragraphProperties17);
                paragraph17.Append(run69);
       
                tableCell12.Append(tableCellProperties12);
                tableCell12.Append(paragraph17);

                TableCell tableCell13 = new TableCell();

                TableCellProperties tableCellProperties13 = new TableCellProperties();
                TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "1008", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders9 = new TableCellBorders();
                TopBorder topBorder10 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder10 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder10 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder10 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

                tableCellBorders9.Append(topBorder10);
                tableCellBorders9.Append(leftBorder10);
                tableCellBorders9.Append(bottomBorder10);
                tableCellBorders9.Append(rightBorder10);
                Shading shading9 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

                tableCellProperties13.Append(tableCellWidth13);
                tableCellProperties13.Append(tableCellBorders9);
                tableCellProperties13.Append(shading9);

                Paragraph paragraph18 = new Paragraph() { RsidParagraphMarkRevision = "001555D6", RsidParagraphAddition = "00831455", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00CC6D34" };

                ParagraphProperties paragraphProperties18 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId18 = new ParagraphStyleId() { Val = "a3" };

                paragraphProperties18.Append(paragraphStyleId18);

                Run run73 = new Run();
                Text text73 = new Text();
                text73.Text = orderi.items[i].Quantity.ToString();

                run73.Append(text73);

                paragraph18.Append(paragraphProperties18);
                paragraph18.Append(run73);

                tableCell13.Append(tableCellProperties13);
                tableCell13.Append(paragraph18);

                TableCell tableCell14 = new TableCell();

                TableCellProperties tableCellProperties14 = new TableCellProperties();
                TableCellWidth tableCellWidth14 = new TableCellWidth() { Width = "1260", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders10 = new TableCellBorders();
                TopBorder topBorder11 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder11 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder11 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder11 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

                tableCellBorders10.Append(topBorder11);
                tableCellBorders10.Append(leftBorder11);
                tableCellBorders10.Append(bottomBorder11);
                tableCellBorders10.Append(rightBorder11);
                Shading shading10 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

                tableCellProperties14.Append(tableCellWidth14);
                tableCellProperties14.Append(tableCellBorders10);
                tableCellProperties14.Append(shading10);

                Paragraph paragraph19 = new Paragraph() { RsidParagraphMarkRevision = "002D5ECD", RsidParagraphAddition = "00831455", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "002D5ECD" };

                ParagraphProperties paragraphProperties19 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId19 = new ParagraphStyleId() { Val = "a3" };

                ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                Languages languages13 = new Languages() { Val = "en-US" };

                paragraphMarkRunProperties3.Append(languages13);

                paragraphProperties19.Append(paragraphStyleId19);
                paragraphProperties19.Append(paragraphMarkRunProperties3);

                Run run74 = new Run();

                RunProperties runProperties11 = new RunProperties();
                Languages languages14 = new Languages() { Val = "en-US" };

                runProperties11.Append(languages14);
                Text text74 = new Text();
                text74.Text = "0";

                run74.Append(runProperties11);
                run74.Append(text74);

                paragraph19.Append(paragraphProperties19);
                paragraph19.Append(run74);

                tableCell14.Append(tableCellProperties14);
                tableCell14.Append(paragraph19);

                TableCell tableCell15 = new TableCell();

                TableCellProperties tableCellProperties15 = new TableCellProperties();
                TableCellWidth tableCellWidth15 = new TableCellWidth() { Width = "2126", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders11 = new TableCellBorders();
                TopBorder topBorder12 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder12 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder12 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder12 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

                tableCellBorders11.Append(topBorder12);
                tableCellBorders11.Append(leftBorder12);
                tableCellBorders11.Append(bottomBorder12);
                tableCellBorders11.Append(rightBorder12);
                Shading shading11 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

                tableCellProperties15.Append(tableCellWidth15);
                tableCellProperties15.Append(tableCellBorders11);
                tableCellProperties15.Append(shading11);

                Paragraph paragraph20 = new Paragraph() { RsidParagraphMarkRevision = "00E86ABA", RsidParagraphAddition = "00831455", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00E86ABA" };

                ParagraphProperties paragraphProperties20 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId20 = new ParagraphStyleId() { Val = "a3" };

                paragraphProperties20.Append(paragraphStyleId20);

                Run run75 = new Run() { RsidRunProperties = "00E86ABA" };

                RunProperties runProperties12 = new RunProperties();
                RunStyle runStyle1 = new RunStyle() { Val = "rvts8" };
                FontSize fontSize1 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };

                runProperties12.Append(runStyle1);
                runProperties12.Append(fontSize1);
                runProperties12.Append(fontSizeComplexScript1);
                Text text75 = new Text();
                text75.Text = orderi.Comment;//////////////

                run75.Append(runProperties12);
                run75.Append(text75);

                paragraph20.Append(paragraphProperties20);
                paragraph20.Append(run75);

                tableCell15.Append(tableCellProperties15);
                tableCell15.Append(paragraph20);

                TableCell tableCell16 = new TableCell();

                TableCellProperties tableCellProperties16 = new TableCellProperties();
                TableCellWidth tableCellWidth16 = new TableCellWidth() { Width = "3176", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders12 = new TableCellBorders();
                TopBorder topBorder13 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder13 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder13 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder13 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

                tableCellBorders12.Append(topBorder13);
                tableCellBorders12.Append(leftBorder13);
                tableCellBorders12.Append(bottomBorder13);
                tableCellBorders12.Append(rightBorder13);
                Shading shading12 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

                tableCellProperties16.Append(tableCellWidth16);
                tableCellProperties16.Append(tableCellBorders12);
                tableCellProperties16.Append(shading12);

                Paragraph paragraph21 = new Paragraph() { RsidParagraphAddition = "00F4161B", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00F4161B" };

                ParagraphProperties paragraphProperties21 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId21 = new ParagraphStyleId() { Val = "a3" };

                paragraphProperties21.Append(paragraphStyleId21);

                Run run76 = new Run();
                Text text76 = new Text();
                text76.Text = orderi.Addrfact;

                run76.Append(text76);

                paragraph21.Append(paragraphProperties21);
                paragraph21.Append(run76);

                Paragraph paragraph22 = new Paragraph() { RsidParagraphAddition = "00F4161B", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00F4161B" };

                ParagraphProperties paragraphProperties22 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId22 = new ParagraphStyleId() { Val = "a3" };

                paragraphProperties22.Append(paragraphStyleId22);

                Run run77 = new Run();
                Text text77 = new Text();
                text77.Text = "Тел. " + orderi.Phonedir;

                run77.Append(text77);

                paragraph22.Append(paragraphProperties22);
                paragraph22.Append(run77);



                Paragraph paragraph24 = new Paragraph() { RsidParagraphAddition = "00F4161B", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00F4161B" };

                ParagraphProperties paragraphProperties24 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId24 = new ParagraphStyleId() { Val = "a3" };

                paragraphProperties24.Append(paragraphStyleId24);

                Run run79 = new Run();
                Text text79 = new Text();
                text79.Text = "ИНН/КПП " + orderi.INN + "/" + orderi.KPP;///////////////////

                run79.Append(text79);

                paragraph24.Append(paragraphProperties24);
                paragraph24.Append(run79);


                Paragraph paragraph26 = new Paragraph() { RsidParagraphAddition = "00F4161B", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00F4161B" };

                ParagraphProperties paragraphProperties26 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId26 = new ParagraphStyleId() { Val = "a3" };

                paragraphProperties26.Append(paragraphStyleId26);

                Run run81 = new Run();
                Text text81 = new Text();
                text81.Text = "Расчётный счёт " + orderi.Schet + " в " + orderi.Bankname;////////////////

                run81.Append(text81);

                paragraph26.Append(paragraphProperties26);
                paragraph26.Append(run81);

                Paragraph paragraph27 = new Paragraph() { RsidParagraphAddition = "00F4161B", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00F4161B" };

                ParagraphProperties paragraphProperties27 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId27 = new ParagraphStyleId() { Val = "a3" };

                paragraphProperties27.Append(paragraphStyleId27);

                Run run82 = new Run();
                Text text82 = new Text();
                text82.Text = "БИК " + orderi.BIK;

                run82.Append(text82);

                paragraph27.Append(paragraphProperties27);
                paragraph27.Append(run82);

                Paragraph paragraph28 = new Paragraph() { RsidParagraphAddition = "00F4161B", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00F4161B" };

                ParagraphProperties paragraphProperties28 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId28 = new ParagraphStyleId() { Val = "a3" };

                paragraphProperties28.Append(paragraphStyleId28);

                Run run83 = new Run();
                Text text83 = new Text();
                text83.Text = orderi.Positiondir + " " + orderi.FIOdir;////////////////////

                run83.Append(text83);

                paragraph28.Append(paragraphProperties28);
                paragraph28.Append(run83);

                Paragraph paragraph29 = new Paragraph() { RsidParagraphAddition = "00F4161B", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00F4161B" };

                ParagraphProperties paragraphProperties29 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId29 = new ParagraphStyleId() { Val = "a3" };

                paragraphProperties29.Append(paragraphStyleId29);

                Run run84 = new Run();
                Text text84 = new Text();
                text84.Text = orderi.Organizationshort;////////////////
                run84.Append(text84);

                paragraph29.Append(paragraphProperties29);
                paragraph29.Append(run84);

                Paragraph paragraph30 = new Paragraph() { RsidParagraphMarkRevision = "00A16BF1", RsidParagraphAddition = "00831455", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00F4161B" };

                ParagraphProperties paragraphProperties30 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId30 = new ParagraphStyleId() { Val = "a3" };

                paragraphProperties30.Append(paragraphStyleId30);

                Run run85 = new Run();
                Text text85 = new Text();
                text85.Text = "";

                run85.Append(text85);

                paragraph30.Append(paragraphProperties30);
                paragraph30.Append(run85);

                tableCell16.Append(tableCellProperties16);
                tableCell16.Append(paragraph21);
                tableCell16.Append(paragraph22);

                tableCell16.Append(paragraph24);

                tableCell16.Append(paragraph26);
                tableCell16.Append(paragraph27);
                tableCell16.Append(paragraph28);
                tableCell16.Append(paragraph29);
                tableCell16.Append(paragraph30);

                tableRow6.Append(tablePropertyExceptions2);
                tableRow6.Append(tableRowProperties2);
                tableRow6.Append(tableCell11);
                tableRow6.Append(tableCell12);
                tableRow6.Append(tableCell13);
                tableRow6.Append(tableCell14);
                tableRow6.Append(tableCell15);
                tableRow6.Append(tableCell16);
         
                table2.Append(tableRow6);
            }
 ////////////////////////////////////////////////////////////////////////////////////////]
            //////////////////////////////////////////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////.



            Paragraph paragraph37 = new Paragraph() { RsidParagraphMarkRevision = "001555D6", RsidParagraphAddition = "00831455", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "00831455" };

            ParagraphProperties paragraphProperties37 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId37 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties37.Append(paragraphStyleId37);

            paragraph37.Append(paragraphProperties37);

            Paragraph paragraph38 = new Paragraph() { RsidParagraphAddition = "00831455", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "006E1D99" };

            ParagraphProperties paragraphProperties38 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId38 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties38.Append(paragraphStyleId38);

            Run run86 = new Run() { RsidRunProperties = "006E1D99" };

            RunProperties runProperties13 = new RunProperties();
            Color color1 = new Color() { Val = "FF0000" };

            runProperties13.Append(color1);
            Text text86 = new Text();
            text86.Text = "Примечание:";

            run86.Append(runProperties13);
            run86.Append(text86);

            Run run87 = new Run();
            Text text87 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text87.Text = " Заявка заполняется и высылается в ОАО «";

            run87.Append(text87);
            ProofError proofError13 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run88 = new Run();
            Text text88 = new Text();
            text88.Text = "ИнфоТеКС";

            run88.Append(text88);
            ProofError proofError14 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run89 = new Run();
            Text text89 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text89.Text = "» в том формате как она есть, без подписей и печатей, на адрес ";

            run89.Append(text89);

            Hyperlink hyperlink1 = new Hyperlink() { History = true, Id = "rId8" };

            Run run90 = new Run() { RsidRunProperties = "00D7526F" };

            RunProperties runProperties14 = new RunProperties();
            RunStyle runStyle2 = new RunStyle() { Val = "a7" };
            Languages languages18 = new Languages() { Val = "en-US" };

            runProperties14.Append(runStyle2);
            runProperties14.Append(languages18);
            Text text90 = new Text();
            text90.Text = "Soft";

            run90.Append(runProperties14);
            run90.Append(text90);

            Run run91 = new Run() { RsidRunProperties = "00D7526F" };

            RunProperties runProperties15 = new RunProperties();
            RunStyle runStyle3 = new RunStyle() { Val = "a7" };

            runProperties15.Append(runStyle3);
            Text text91 = new Text();
            text91.Text = "@";

            run91.Append(runProperties15);
            run91.Append(text91);
            ProofError proofError15 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run92 = new Run() { RsidRunProperties = "00D7526F" };

            RunProperties runProperties16 = new RunProperties();
            RunStyle runStyle4 = new RunStyle() { Val = "a7" };
            Languages languages19 = new Languages() { Val = "en-US" };

            runProperties16.Append(runStyle4);
            runProperties16.Append(languages19);
            Text text92 = new Text();
            text92.Text = "infotecs";

            run92.Append(runProperties16);
            run92.Append(text92);
            ProofError proofError16 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run93 = new Run() { RsidRunProperties = "00D7526F" };

            RunProperties runProperties17 = new RunProperties();
            RunStyle runStyle5 = new RunStyle() { Val = "a7" };

            runProperties17.Append(runStyle5);
            Text text93 = new Text();
            text93.Text = ".";

            run93.Append(runProperties17);
            run93.Append(text93);
            ProofError proofError17 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run94 = new Run() { RsidRunProperties = "00D7526F" };

            RunProperties runProperties18 = new RunProperties();
            RunStyle runStyle6 = new RunStyle() { Val = "a7" };
            Languages languages20 = new Languages() { Val = "en-US" };

            runProperties18.Append(runStyle6);
            runProperties18.Append(languages20);
            Text text94 = new Text();
            text94.Text = "ru";

            run94.Append(runProperties18);
            run94.Append(text94);
            ProofError proofError18 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            hyperlink1.Append(run90);
            hyperlink1.Append(run91);
            hyperlink1.Append(proofError15);
            hyperlink1.Append(run92);
            hyperlink1.Append(proofError16);
            hyperlink1.Append(run93);
            hyperlink1.Append(proofError17);
            hyperlink1.Append(run94);
            hyperlink1.Append(proofError18);

            paragraph38.Append(paragraphProperties38);
            paragraph38.Append(run86);
            paragraph38.Append(run87);
            paragraph38.Append(proofError13);
            paragraph38.Append(run88);
            paragraph38.Append(proofError14);
            paragraph38.Append(run89);
            paragraph38.Append(hyperlink1);

            Paragraph paragraph39 = new Paragraph() { RsidParagraphMarkRevision = "006E1D99", RsidParagraphAddition = "008D46DE", RsidParagraphProperties = "00D37E6B", RsidRunAdditionDefault = "008D46DE" };

            ParagraphProperties paragraphProperties39 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId39 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties39.Append(paragraphStyleId39);

            Run run95 = new Run();
            Text text95 = new Text();
            text95.Text = "Все поля обязательны к заполнению.";

            run95.Append(text95);

            paragraph39.Append(paragraphProperties39);
            paragraph39.Append(run95);

            SectionProperties sectionProperties1 = new SectionProperties() { RsidRPr = "006E1D99", RsidR = "008D46DE", RsidSect = "00831455" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)11906U, Height = (UInt32Value)16838U };
            PageMargin pageMargin1 = new PageMargin() { Top = 567, Right = (UInt32Value)850U, Bottom = 1134, Left = (UInt32Value)1701U, Header = (UInt32Value)708U, Footer = (UInt32Value)708U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "708" };
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            body1.Append(paragraph1);
            body1.Append(paragraph2);
            body1.Append(table1);
            body1.Append(paragraph9);
            body1.Append(table2);
            body1.Append(paragraph37);
            body1.Append(paragraph38);
            body1.Append(paragraph39);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
        }

        // Generates content of stylesWithEffectsPart1.
        private void GenerateStylesWithEffectsPart1Content(StylesWithEffectsPart stylesWithEffectsPart1)
        {
            Styles styles1 = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            styles1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            styles1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            styles1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            styles1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            styles1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            styles1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            styles1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            styles1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            styles1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            styles1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            styles1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            styles1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts5 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
            Languages languages21 = new Languages() { Val = "ru-RU", EastAsia = "ru-RU", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts5);
            runPropertiesBaseStyle1.Append(languages21);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);
            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 99, DefaultSemiHidden = true, DefaultUnhideWhenUsed = true, DefaultPrimaryStyle = false, Count = 267 };
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo() { Name = "Normal", UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo() { Name = "heading 1", UiPriority = 9, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo() { Name = "heading 2", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo() { Name = "heading 3", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo() { Name = "heading 4", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo() { Name = "heading 5", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo() { Name = "heading 6", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo() { Name = "heading 7", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo() { Name = "heading 8", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo() { Name = "heading 9", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo() { Name = "toc 1", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo() { Name = "toc 2", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo() { Name = "toc 3", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo() { Name = "toc 4", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo() { Name = "toc 5", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo() { Name = "toc 6", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo() { Name = "toc 7", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo() { Name = "toc 8", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo() { Name = "toc 9", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo() { Name = "caption", UiPriority = 35, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo() { Name = "Title", UiPriority = 10, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo() { Name = "Default Paragraph Font", UiPriority = 1 };
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo() { Name = "Subtitle", UiPriority = 11, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo() { Name = "Hyperlink", UiPriority = 0 };
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo() { Name = "Strong", UiPriority = 22, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo() { Name = "Emphasis", UiPriority = 20, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo() { Name = "Table Grid", UiPriority = 59, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo() { Name = "Revision", UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37 };
            LatentStyleExceptionInfo latentStyleExceptionInfo138 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, PrimaryStyle = true };

            latentStyles1.Append(latentStyleExceptionInfo1);
            latentStyles1.Append(latentStyleExceptionInfo2);
            latentStyles1.Append(latentStyleExceptionInfo3);
            latentStyles1.Append(latentStyleExceptionInfo4);
            latentStyles1.Append(latentStyleExceptionInfo5);
            latentStyles1.Append(latentStyleExceptionInfo6);
            latentStyles1.Append(latentStyleExceptionInfo7);
            latentStyles1.Append(latentStyleExceptionInfo8);
            latentStyles1.Append(latentStyleExceptionInfo9);
            latentStyles1.Append(latentStyleExceptionInfo10);
            latentStyles1.Append(latentStyleExceptionInfo11);
            latentStyles1.Append(latentStyleExceptionInfo12);
            latentStyles1.Append(latentStyleExceptionInfo13);
            latentStyles1.Append(latentStyleExceptionInfo14);
            latentStyles1.Append(latentStyleExceptionInfo15);
            latentStyles1.Append(latentStyleExceptionInfo16);
            latentStyles1.Append(latentStyleExceptionInfo17);
            latentStyles1.Append(latentStyleExceptionInfo18);
            latentStyles1.Append(latentStyleExceptionInfo19);
            latentStyles1.Append(latentStyleExceptionInfo20);
            latentStyles1.Append(latentStyleExceptionInfo21);
            latentStyles1.Append(latentStyleExceptionInfo22);
            latentStyles1.Append(latentStyleExceptionInfo23);
            latentStyles1.Append(latentStyleExceptionInfo24);
            latentStyles1.Append(latentStyleExceptionInfo25);
            latentStyles1.Append(latentStyleExceptionInfo26);
            latentStyles1.Append(latentStyleExceptionInfo27);
            latentStyles1.Append(latentStyleExceptionInfo28);
            latentStyles1.Append(latentStyleExceptionInfo29);
            latentStyles1.Append(latentStyleExceptionInfo30);
            latentStyles1.Append(latentStyleExceptionInfo31);
            latentStyles1.Append(latentStyleExceptionInfo32);
            latentStyles1.Append(latentStyleExceptionInfo33);
            latentStyles1.Append(latentStyleExceptionInfo34);
            latentStyles1.Append(latentStyleExceptionInfo35);
            latentStyles1.Append(latentStyleExceptionInfo36);
            latentStyles1.Append(latentStyleExceptionInfo37);
            latentStyles1.Append(latentStyleExceptionInfo38);
            latentStyles1.Append(latentStyleExceptionInfo39);
            latentStyles1.Append(latentStyleExceptionInfo40);
            latentStyles1.Append(latentStyleExceptionInfo41);
            latentStyles1.Append(latentStyleExceptionInfo42);
            latentStyles1.Append(latentStyleExceptionInfo43);
            latentStyles1.Append(latentStyleExceptionInfo44);
            latentStyles1.Append(latentStyleExceptionInfo45);
            latentStyles1.Append(latentStyleExceptionInfo46);
            latentStyles1.Append(latentStyleExceptionInfo47);
            latentStyles1.Append(latentStyleExceptionInfo48);
            latentStyles1.Append(latentStyleExceptionInfo49);
            latentStyles1.Append(latentStyleExceptionInfo50);
            latentStyles1.Append(latentStyleExceptionInfo51);
            latentStyles1.Append(latentStyleExceptionInfo52);
            latentStyles1.Append(latentStyleExceptionInfo53);
            latentStyles1.Append(latentStyleExceptionInfo54);
            latentStyles1.Append(latentStyleExceptionInfo55);
            latentStyles1.Append(latentStyleExceptionInfo56);
            latentStyles1.Append(latentStyleExceptionInfo57);
            latentStyles1.Append(latentStyleExceptionInfo58);
            latentStyles1.Append(latentStyleExceptionInfo59);
            latentStyles1.Append(latentStyleExceptionInfo60);
            latentStyles1.Append(latentStyleExceptionInfo61);
            latentStyles1.Append(latentStyleExceptionInfo62);
            latentStyles1.Append(latentStyleExceptionInfo63);
            latentStyles1.Append(latentStyleExceptionInfo64);
            latentStyles1.Append(latentStyleExceptionInfo65);
            latentStyles1.Append(latentStyleExceptionInfo66);
            latentStyles1.Append(latentStyleExceptionInfo67);
            latentStyles1.Append(latentStyleExceptionInfo68);
            latentStyles1.Append(latentStyleExceptionInfo69);
            latentStyles1.Append(latentStyleExceptionInfo70);
            latentStyles1.Append(latentStyleExceptionInfo71);
            latentStyles1.Append(latentStyleExceptionInfo72);
            latentStyles1.Append(latentStyleExceptionInfo73);
            latentStyles1.Append(latentStyleExceptionInfo74);
            latentStyles1.Append(latentStyleExceptionInfo75);
            latentStyles1.Append(latentStyleExceptionInfo76);
            latentStyles1.Append(latentStyleExceptionInfo77);
            latentStyles1.Append(latentStyleExceptionInfo78);
            latentStyles1.Append(latentStyleExceptionInfo79);
            latentStyles1.Append(latentStyleExceptionInfo80);
            latentStyles1.Append(latentStyleExceptionInfo81);
            latentStyles1.Append(latentStyleExceptionInfo82);
            latentStyles1.Append(latentStyleExceptionInfo83);
            latentStyles1.Append(latentStyleExceptionInfo84);
            latentStyles1.Append(latentStyleExceptionInfo85);
            latentStyles1.Append(latentStyleExceptionInfo86);
            latentStyles1.Append(latentStyleExceptionInfo87);
            latentStyles1.Append(latentStyleExceptionInfo88);
            latentStyles1.Append(latentStyleExceptionInfo89);
            latentStyles1.Append(latentStyleExceptionInfo90);
            latentStyles1.Append(latentStyleExceptionInfo91);
            latentStyles1.Append(latentStyleExceptionInfo92);
            latentStyles1.Append(latentStyleExceptionInfo93);
            latentStyles1.Append(latentStyleExceptionInfo94);
            latentStyles1.Append(latentStyleExceptionInfo95);
            latentStyles1.Append(latentStyleExceptionInfo96);
            latentStyles1.Append(latentStyleExceptionInfo97);
            latentStyles1.Append(latentStyleExceptionInfo98);
            latentStyles1.Append(latentStyleExceptionInfo99);
            latentStyles1.Append(latentStyleExceptionInfo100);
            latentStyles1.Append(latentStyleExceptionInfo101);
            latentStyles1.Append(latentStyleExceptionInfo102);
            latentStyles1.Append(latentStyleExceptionInfo103);
            latentStyles1.Append(latentStyleExceptionInfo104);
            latentStyles1.Append(latentStyleExceptionInfo105);
            latentStyles1.Append(latentStyleExceptionInfo106);
            latentStyles1.Append(latentStyleExceptionInfo107);
            latentStyles1.Append(latentStyleExceptionInfo108);
            latentStyles1.Append(latentStyleExceptionInfo109);
            latentStyles1.Append(latentStyleExceptionInfo110);
            latentStyles1.Append(latentStyleExceptionInfo111);
            latentStyles1.Append(latentStyleExceptionInfo112);
            latentStyles1.Append(latentStyleExceptionInfo113);
            latentStyles1.Append(latentStyleExceptionInfo114);
            latentStyles1.Append(latentStyleExceptionInfo115);
            latentStyles1.Append(latentStyleExceptionInfo116);
            latentStyles1.Append(latentStyleExceptionInfo117);
            latentStyles1.Append(latentStyleExceptionInfo118);
            latentStyles1.Append(latentStyleExceptionInfo119);
            latentStyles1.Append(latentStyleExceptionInfo120);
            latentStyles1.Append(latentStyleExceptionInfo121);
            latentStyles1.Append(latentStyleExceptionInfo122);
            latentStyles1.Append(latentStyleExceptionInfo123);
            latentStyles1.Append(latentStyleExceptionInfo124);
            latentStyles1.Append(latentStyleExceptionInfo125);
            latentStyles1.Append(latentStyleExceptionInfo126);
            latentStyles1.Append(latentStyleExceptionInfo127);
            latentStyles1.Append(latentStyleExceptionInfo128);
            latentStyles1.Append(latentStyleExceptionInfo129);
            latentStyles1.Append(latentStyleExceptionInfo130);
            latentStyles1.Append(latentStyleExceptionInfo131);
            latentStyles1.Append(latentStyleExceptionInfo132);
            latentStyles1.Append(latentStyleExceptionInfo133);
            latentStyles1.Append(latentStyleExceptionInfo134);
            latentStyles1.Append(latentStyleExceptionInfo135);
            latentStyles1.Append(latentStyleExceptionInfo136);
            latentStyles1.Append(latentStyleExceptionInfo137);
            latentStyles1.Append(latentStyleExceptionInfo138);

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "a", Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            Rsid rsid1 = new Rsid() { Val = "00831455" };

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts6 = new RunFonts() { Ascii = "Arial Unicode MS", HighAnsi = "Arial Unicode MS", EastAsia = "Arial Unicode MS", ComplexScript = "Arial Unicode MS" };
            Color color2 = new Color() { Val = "000000" };
            FontSize fontSize2 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "24" };
            Languages languages22 = new Languages() { Val = "ru" };

            styleRunProperties1.Append(runFonts6);
            styleRunProperties1.Append(color2);
            styleRunProperties1.Append(fontSize2);
            styleRunProperties1.Append(fontSizeComplexScript2);
            styleRunProperties1.Append(languages22);

            style1.Append(styleName1);
            style1.Append(rsid1);
            style1.Append(styleRunProperties1);

            Style style2 = new Style() { Type = StyleValues.Paragraph, StyleId = "1" };
            StyleName styleName2 = new StyleName() { Val = "heading 1" };
            BasedOn basedOn1 = new BasedOn() { Val = "a" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "a" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "10" };
            UIPriority uIPriority1 = new UIPriority() { Val = 9 };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            Rsid rsid2 = new Rsid() { Val = "00D37E6B" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "240", After = "60" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 0 };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            RunFonts runFonts7 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            Kern kern1 = new Kern() { Val = (UInt32Value)32U };
            FontSize fontSize3 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties2.Append(runFonts7);
            styleRunProperties2.Append(bold1);
            styleRunProperties2.Append(boldComplexScript1);
            styleRunProperties2.Append(kern1);
            styleRunProperties2.Append(fontSize3);
            styleRunProperties2.Append(fontSizeComplexScript3);

            style2.Append(styleName2);
            style2.Append(basedOn1);
            style2.Append(nextParagraphStyle1);
            style2.Append(linkedStyle1);
            style2.Append(uIPriority1);
            style2.Append(primaryStyle1);
            style2.Append(rsid2);
            style2.Append(styleParagraphProperties1);
            style2.Append(styleRunProperties2);

            Style style3 = new Style() { Type = StyleValues.Character, StyleId = "a0", Default = true };
            StyleName styleName3 = new StyleName() { Val = "Default Paragraph Font" };
            UIPriority uIPriority2 = new UIPriority() { Val = 1 };
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();

            style3.Append(styleName3);
            style3.Append(uIPriority2);
            style3.Append(unhideWhenUsed1);

            Style style4 = new Style() { Type = StyleValues.Table, StyleId = "a1", Default = true };
            StyleName styleName4 = new StyleName() { Val = "Normal Table" };
            UIPriority uIPriority3 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle2 = new PrimaryStyle();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation3 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault5 = new TableCellMarginDefault();
            TopMargin topMargin4 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin2 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin4 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin2 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault5.Append(topMargin4);
            tableCellMarginDefault5.Append(tableCellLeftMargin2);
            tableCellMarginDefault5.Append(bottomMargin4);
            tableCellMarginDefault5.Append(tableCellRightMargin2);

            styleTableProperties1.Append(tableIndentation3);
            styleTableProperties1.Append(tableCellMarginDefault5);

            style4.Append(styleName4);
            style4.Append(uIPriority3);
            style4.Append(semiHidden1);
            style4.Append(unhideWhenUsed2);
            style4.Append(primaryStyle2);
            style4.Append(styleTableProperties1);

            Style style5 = new Style() { Type = StyleValues.Numbering, StyleId = "a2", Default = true };
            StyleName styleName5 = new StyleName() { Val = "No List" };
            UIPriority uIPriority4 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden2 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed3 = new UnhideWhenUsed();

            style5.Append(styleName5);
            style5.Append(uIPriority4);
            style5.Append(semiHidden2);
            style5.Append(unhideWhenUsed3);

            Style style6 = new Style() { Type = StyleValues.Paragraph, StyleId = "a3" };
            StyleName styleName6 = new StyleName() { Val = "No Spacing" };
            UIPriority uIPriority5 = new UIPriority() { Val = 1 };
            PrimaryStyle primaryStyle3 = new PrimaryStyle();
            Rsid rsid3 = new Rsid() { Val = "00D37E6B" };

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            RunFonts runFonts8 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            FontSize fontSize4 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "22" };
            Languages languages23 = new Languages() { EastAsia = "en-US" };

            styleRunProperties3.Append(runFonts8);
            styleRunProperties3.Append(fontSize4);
            styleRunProperties3.Append(fontSizeComplexScript4);
            styleRunProperties3.Append(languages23);

            style6.Append(styleName6);
            style6.Append(uIPriority5);
            style6.Append(primaryStyle3);
            style6.Append(rsid3);
            style6.Append(styleRunProperties3);

            Style style7 = new Style() { Type = StyleValues.Character, StyleId = "a4", CustomStyle = true };
            StyleName styleName7 = new StyleName() { Val = "Сноска_" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "a5" };
            Rsid rsid4 = new Rsid() { Val = "00831455" };

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            RunFonts runFonts9 = new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI", EastAsia = "Segoe UI", ComplexScript = "Segoe UI" };
            FontSize fontSize5 = new FontSize() { Val = "19" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "19" };
            Shading shading19 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            styleRunProperties4.Append(runFonts9);
            styleRunProperties4.Append(fontSize5);
            styleRunProperties4.Append(fontSizeComplexScript5);
            styleRunProperties4.Append(shading19);

            style7.Append(styleName7);
            style7.Append(linkedStyle2);
            style7.Append(rsid4);
            style7.Append(styleRunProperties4);

            Style style8 = new Style() { Type = StyleValues.Character, StyleId = "a6", CustomStyle = true };
            StyleName styleName8 = new StyleName() { Val = "Основной текст_" };
            LinkedStyle linkedStyle3 = new LinkedStyle() { Val = "11" };
            Rsid rsid5 = new Rsid() { Val = "00831455" };

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI", EastAsia = "Segoe UI", ComplexScript = "Segoe UI" };
            FontSize fontSize6 = new FontSize() { Val = "19" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "19" };
            Shading shading20 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            styleRunProperties5.Append(runFonts10);
            styleRunProperties5.Append(fontSize6);
            styleRunProperties5.Append(fontSizeComplexScript6);
            styleRunProperties5.Append(shading20);

            style8.Append(styleName8);
            style8.Append(linkedStyle3);
            style8.Append(rsid5);
            style8.Append(styleRunProperties5);

            Style style9 = new Style() { Type = StyleValues.Paragraph, StyleId = "a5", CustomStyle = true };
            StyleName styleName9 = new StyleName() { Val = "Сноска" };
            BasedOn basedOn2 = new BasedOn() { Val = "a" };
            LinkedStyle linkedStyle4 = new LinkedStyle() { Val = "a4" };
            Rsid rsid6 = new Rsid() { Val = "00831455" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
            Shading shading21 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Line = "0", LineRule = LineSpacingRuleValues.AtLeast };

            styleParagraphProperties2.Append(shading21);
            styleParagraphProperties2.Append(spacingBetweenLines2);

            StyleRunProperties styleRunProperties6 = new StyleRunProperties();
            RunFonts runFonts11 = new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI", EastAsia = "Segoe UI", ComplexScript = "Times New Roman" };
            Color color3 = new Color() { Val = "auto" };
            FontSize fontSize7 = new FontSize() { Val = "19" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "19" };
            Languages languages24 = new Languages() { Val = "x-none", EastAsia = "x-none" };

            styleRunProperties6.Append(runFonts11);
            styleRunProperties6.Append(color3);
            styleRunProperties6.Append(fontSize7);
            styleRunProperties6.Append(fontSizeComplexScript7);
            styleRunProperties6.Append(languages24);

            style9.Append(styleName9);
            style9.Append(basedOn2);
            style9.Append(linkedStyle4);
            style9.Append(rsid6);
            style9.Append(styleParagraphProperties2);
            style9.Append(styleRunProperties6);

            Style style10 = new Style() { Type = StyleValues.Paragraph, StyleId = "11", CustomStyle = true };
            StyleName styleName10 = new StyleName() { Val = "Основной текст1" };
            BasedOn basedOn3 = new BasedOn() { Val = "a" };
            LinkedStyle linkedStyle5 = new LinkedStyle() { Val = "a6" };
            Rsid rsid7 = new Rsid() { Val = "00831455" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();
            Shading shading22 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Line = "0", LineRule = LineSpacingRuleValues.AtLeast };

            styleParagraphProperties3.Append(shading22);
            styleParagraphProperties3.Append(spacingBetweenLines3);

            StyleRunProperties styleRunProperties7 = new StyleRunProperties();
            RunFonts runFonts12 = new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI", EastAsia = "Segoe UI", ComplexScript = "Times New Roman" };
            Color color4 = new Color() { Val = "auto" };
            FontSize fontSize8 = new FontSize() { Val = "19" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "19" };
            Languages languages25 = new Languages() { Val = "x-none", EastAsia = "x-none" };

            styleRunProperties7.Append(runFonts12);
            styleRunProperties7.Append(color4);
            styleRunProperties7.Append(fontSize8);
            styleRunProperties7.Append(fontSizeComplexScript8);
            styleRunProperties7.Append(languages25);

            style10.Append(styleName10);
            style10.Append(basedOn3);
            style10.Append(linkedStyle5);
            style10.Append(rsid7);
            style10.Append(styleParagraphProperties3);
            style10.Append(styleRunProperties7);

            Style style11 = new Style() { Type = StyleValues.Paragraph, StyleId = "2", CustomStyle = true };
            StyleName styleName11 = new StyleName() { Val = "Знак2" };
            BasedOn basedOn4 = new BasedOn() { Val = "a" };
            Rsid rsid8 = new Rsid() { Val = "00831455" };

            StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { After = "160", Line = "240", LineRule = LineSpacingRuleValues.Exact };

            styleParagraphProperties4.Append(spacingBetweenLines4);

            StyleRunProperties styleRunProperties8 = new StyleRunProperties();
            RunFonts runFonts13 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana", EastAsia = "Times New Roman", ComplexScript = "Verdana" };
            Color color5 = new Color() { Val = "auto" };
            FontSize fontSize9 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "20" };
            Languages languages26 = new Languages() { Val = "en-US", EastAsia = "en-US" };

            styleRunProperties8.Append(runFonts13);
            styleRunProperties8.Append(color5);
            styleRunProperties8.Append(fontSize9);
            styleRunProperties8.Append(fontSizeComplexScript9);
            styleRunProperties8.Append(languages26);

            style11.Append(styleName11);
            style11.Append(basedOn4);
            style11.Append(rsid8);
            style11.Append(styleParagraphProperties4);
            style11.Append(styleRunProperties8);

            Style style12 = new Style() { Type = StyleValues.Character, StyleId = "a7" };
            StyleName styleName12 = new StyleName() { Val = "Hyperlink" };
            Rsid rsid9 = new Rsid() { Val = "00831455" };

            StyleRunProperties styleRunProperties9 = new StyleRunProperties();
            Color color6 = new Color() { Val = "0000FF" };
            Underline underline1 = new Underline() { Val = UnderlineValues.Single };

            styleRunProperties9.Append(color6);
            styleRunProperties9.Append(underline1);

            style12.Append(styleName12);
            style12.Append(rsid9);
            style12.Append(styleRunProperties9);

            Style style13 = new Style() { Type = StyleValues.Paragraph, StyleId = "a8" };
            StyleName styleName13 = new StyleName() { Val = "List Paragraph" };
            BasedOn basedOn5 = new BasedOn() { Val = "a" };
            UIPriority uIPriority6 = new UIPriority() { Val = 34 };
            PrimaryStyle primaryStyle4 = new PrimaryStyle();
            Rsid rsid10 = new Rsid() { Val = "00FE63E5" };

            StyleParagraphProperties styleParagraphProperties5 = new StyleParagraphProperties();
            Indentation indentation1 = new Indentation() { Left = "720" };
            ContextualSpacing contextualSpacing1 = new ContextualSpacing();

            styleParagraphProperties5.Append(indentation1);
            styleParagraphProperties5.Append(contextualSpacing1);

            style13.Append(styleName13);
            style13.Append(basedOn5);
            style13.Append(uIPriority6);
            style13.Append(primaryStyle4);
            style13.Append(rsid10);
            style13.Append(styleParagraphProperties5);

            Style style14 = new Style() { Type = StyleValues.Character, StyleId = "FontStyle77", CustomStyle = true };
            StyleName styleName14 = new StyleName() { Val = "Font Style77" };
            UIPriority uIPriority7 = new UIPriority() { Val = 99 };
            Rsid rsid11 = new Rsid() { Val = "00C52219" };

            StyleRunProperties styleRunProperties10 = new StyleRunProperties();
            RunFonts runFonts14 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            FontSize fontSize10 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties10.Append(runFonts14);
            styleRunProperties10.Append(bold2);
            styleRunProperties10.Append(boldComplexScript2);
            styleRunProperties10.Append(fontSize10);
            styleRunProperties10.Append(fontSizeComplexScript10);

            style14.Append(styleName14);
            style14.Append(uIPriority7);
            style14.Append(rsid11);
            style14.Append(styleRunProperties10);

            Style style15 = new Style() { Type = StyleValues.Table, StyleId = "a9" };
            StyleName styleName15 = new StyleName() { Val = "Table Grid" };
            BasedOn basedOn6 = new BasedOn() { Val = "a1" };
            UIPriority uIPriority8 = new UIPriority() { Val = 59 };
            Rsid rsid12 = new Rsid() { Val = "00F444F2" };

            StyleTableProperties styleTableProperties2 = new StyleTableProperties();

            TableBorders tableBorders2 = new TableBorders();
            TopBorder topBorder20 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder20 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder20 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder20 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder2 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder2 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders2.Append(topBorder20);
            tableBorders2.Append(leftBorder20);
            tableBorders2.Append(bottomBorder20);
            tableBorders2.Append(rightBorder20);
            tableBorders2.Append(insideHorizontalBorder2);
            tableBorders2.Append(insideVerticalBorder2);

            styleTableProperties2.Append(tableBorders2);

            style15.Append(styleName15);
            style15.Append(basedOn6);
            style15.Append(uIPriority8);
            style15.Append(rsid12);
            style15.Append(styleTableProperties2);

            Style style16 = new Style() { Type = StyleValues.Character, StyleId = "rvts8", CustomStyle = true };
            StyleName styleName16 = new StyleName() { Val = "rvts8" };
            BasedOn basedOn7 = new BasedOn() { Val = "a0" };
            Rsid rsid13 = new Rsid() { Val = "00E86ABA" };

            StyleRunProperties styleRunProperties11 = new StyleRunProperties();
            Bold bold3 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();
            FontSize fontSize11 = new FontSize() { Val = "36" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "36" };

            styleRunProperties11.Append(bold3);
            styleRunProperties11.Append(boldComplexScript3);
            styleRunProperties11.Append(fontSize11);
            styleRunProperties11.Append(fontSizeComplexScript11);

            style16.Append(styleName16);
            style16.Append(basedOn7);
            style16.Append(rsid13);
            style16.Append(styleRunProperties11);

            Style style17 = new Style() { Type = StyleValues.Paragraph, StyleId = "aa" };
            StyleName styleName17 = new StyleName() { Val = "Normal (Web)" };
            BasedOn basedOn8 = new BasedOn() { Val = "a" };
            UIPriority uIPriority9 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden3 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed4 = new UnhideWhenUsed();
            Rsid rsid14 = new Rsid() { Val = "00E86ABA" };

            StyleRunProperties styleRunProperties12 = new StyleRunProperties();
            RunFonts runFonts15 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
            Color color7 = new Color() { Val = "auto" };
            Languages languages27 = new Languages() { Val = "ru-RU" };

            styleRunProperties12.Append(runFonts15);
            styleRunProperties12.Append(color7);
            styleRunProperties12.Append(languages27);

            style17.Append(styleName17);
            style17.Append(basedOn8);
            style17.Append(uIPriority9);
            style17.Append(semiHidden3);
            style17.Append(unhideWhenUsed4);
            style17.Append(rsid14);
            style17.Append(styleRunProperties12);

            Style style18 = new Style() { Type = StyleValues.Character, StyleId = "rvts6", CustomStyle = true };
            StyleName styleName18 = new StyleName() { Val = "rvts6" };
            BasedOn basedOn9 = new BasedOn() { Val = "a0" };
            Rsid rsid15 = new Rsid() { Val = "00E86ABA" };

            StyleRunProperties styleRunProperties13 = new StyleRunProperties();
            FontSize fontSize12 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "28" };

            styleRunProperties13.Append(fontSize12);
            styleRunProperties13.Append(fontSizeComplexScript12);

            style18.Append(styleName18);
            style18.Append(basedOn9);
            style18.Append(rsid15);
            style18.Append(styleRunProperties13);

            Style style19 = new Style() { Type = StyleValues.Character, StyleId = "12", CustomStyle = true };
            StyleName styleName19 = new StyleName() { Val = "Заголовок №1" };
            Rsid rsid16 = new Rsid() { Val = "00A16BF1" };

            StyleRunProperties styleRunProperties14 = new StyleRunProperties();
            RunFonts runFonts16 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold4 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();
            Italic italic1 = new Italic() { Val = false };
            ItalicComplexScript italicComplexScript1 = new ItalicComplexScript() { Val = false };
            Caps caps1 = new Caps() { Val = false };
            SmallCaps smallCaps1 = new SmallCaps() { Val = false };
            Strike strike1 = new Strike() { Val = false };
            DoubleStrike doubleStrike1 = new DoubleStrike() { Val = false };
            Color color8 = new Color() { Val = "000000" };
            Spacing spacing1 = new Spacing() { Val = 18 };
            CharacterScale characterScale1 = new CharacterScale() { Val = 100 };
            Position position1 = new Position() { Val = "0" };
            FontSize fontSize13 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "23" };
            Underline underline2 = new Underline() { Val = UnderlineValues.Single };
            VerticalTextAlignment verticalTextAlignment1 = new VerticalTextAlignment() { Val = VerticalPositionValues.Baseline };
            Languages languages28 = new Languages() { Val = "ru-RU" };

            styleRunProperties14.Append(runFonts16);
            styleRunProperties14.Append(bold4);
            styleRunProperties14.Append(boldComplexScript4);
            styleRunProperties14.Append(italic1);
            styleRunProperties14.Append(italicComplexScript1);
            styleRunProperties14.Append(caps1);
            styleRunProperties14.Append(smallCaps1);
            styleRunProperties14.Append(strike1);
            styleRunProperties14.Append(doubleStrike1);
            styleRunProperties14.Append(color8);
            styleRunProperties14.Append(spacing1);
            styleRunProperties14.Append(characterScale1);
            styleRunProperties14.Append(position1);
            styleRunProperties14.Append(fontSize13);
            styleRunProperties14.Append(fontSizeComplexScript13);
            styleRunProperties14.Append(underline2);
            styleRunProperties14.Append(verticalTextAlignment1);
            styleRunProperties14.Append(languages28);

            style19.Append(styleName19);
            style19.Append(rsid16);
            style19.Append(styleRunProperties14);

            Style style20 = new Style() { Type = StyleValues.Paragraph, StyleId = "Default", CustomStyle = true };
            StyleName styleName20 = new StyleName() { Val = "Default" };
            Rsid rsid17 = new Rsid() { Val = "002D5ECD" };

            StyleParagraphProperties styleParagraphProperties6 = new StyleParagraphProperties();
            AutoSpaceDE autoSpaceDE1 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN1 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent1 = new AdjustRightIndent() { Val = false };

            styleParagraphProperties6.Append(autoSpaceDE1);
            styleParagraphProperties6.Append(autoSpaceDN1);
            styleParagraphProperties6.Append(adjustRightIndent1);

            StyleRunProperties styleRunProperties15 = new StyleRunProperties();
            RunFonts runFonts17 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman" };
            Color color9 = new Color() { Val = "000000" };
            FontSize fontSize14 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties15.Append(runFonts17);
            styleRunProperties15.Append(color9);
            styleRunProperties15.Append(fontSize14);
            styleRunProperties15.Append(fontSizeComplexScript14);

            style20.Append(styleName20);
            style20.Append(rsid17);
            style20.Append(styleParagraphProperties6);
            style20.Append(styleRunProperties15);

            Style style21 = new Style() { Type = StyleValues.Character, StyleId = "10", CustomStyle = true };
            StyleName styleName21 = new StyleName() { Val = "Заголовок 1 Знак" };
            BasedOn basedOn10 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle6 = new LinkedStyle() { Val = "1" };
            UIPriority uIPriority10 = new UIPriority() { Val = 9 };
            Rsid rsid18 = new Rsid() { Val = "00D37E6B" };

            StyleRunProperties styleRunProperties16 = new StyleRunProperties();
            RunFonts runFonts18 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold5 = new Bold();
            BoldComplexScript boldComplexScript5 = new BoldComplexScript();
            Color color10 = new Color() { Val = "000000" };
            Kern kern2 = new Kern() { Val = (UInt32Value)32U };
            FontSize fontSize15 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "32" };
            Languages languages29 = new Languages() { Val = "ru" };

            styleRunProperties16.Append(runFonts18);
            styleRunProperties16.Append(bold5);
            styleRunProperties16.Append(boldComplexScript5);
            styleRunProperties16.Append(color10);
            styleRunProperties16.Append(kern2);
            styleRunProperties16.Append(fontSize15);
            styleRunProperties16.Append(fontSizeComplexScript15);
            styleRunProperties16.Append(languages29);

            style21.Append(styleName21);
            style21.Append(basedOn10);
            style21.Append(linkedStyle6);
            style21.Append(uIPriority10);
            style21.Append(rsid18);
            style21.Append(styleRunProperties16);

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);
            styles1.Append(style6);
            styles1.Append(style7);
            styles1.Append(style8);
            styles1.Append(style9);
            styles1.Append(style10);
            styles1.Append(style11);
            styles1.Append(style12);
            styles1.Append(style13);
            styles1.Append(style14);
            styles1.Append(style15);
            styles1.Append(style16);
            styles1.Append(style17);
            styles1.Append(style18);
            styles1.Append(style19);
            styles1.Append(style20);
            styles1.Append(style21);

            stylesWithEffectsPart1.Styles = styles1;
        }

        // Generates content of endnotesPart1.
        private void GenerateEndnotesPart1Content(EndnotesPart endnotesPart1)
        {
            Endnotes endnotes1 = new Endnotes() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            endnotes1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            endnotes1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            endnotes1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            endnotes1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            endnotes1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            endnotes1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            endnotes1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            endnotes1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            endnotes1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            endnotes1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            endnotes1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            endnotes1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            endnotes1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            endnotes1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            endnotes1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Endnote endnote1 = new Endnote() { Type = FootnoteEndnoteValues.Separator, Id = -1 };

            Paragraph paragraph40 = new Paragraph() { RsidParagraphAddition = "00EF038F", RsidParagraphProperties = "00831455", RsidRunAdditionDefault = "00EF038F" };

            Run run96 = new Run();
            SeparatorMark separatorMark1 = new SeparatorMark();

            run96.Append(separatorMark1);

            paragraph40.Append(run96);

            endnote1.Append(paragraph40);

            Endnote endnote2 = new Endnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph41 = new Paragraph() { RsidParagraphAddition = "00EF038F", RsidParagraphProperties = "00831455", RsidRunAdditionDefault = "00EF038F" };

            Run run97 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark1 = new ContinuationSeparatorMark();

            run97.Append(continuationSeparatorMark1);

            paragraph41.Append(run97);

            endnote2.Append(paragraph41);

            endnotes1.Append(endnote1);
            endnotes1.Append(endnote2);

            endnotesPart1.Endnotes = endnotes1;
        }

        // Generates content of styleDefinitionsPart1.
        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles2 = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            styles2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            styles2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            styles2.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");

            DocDefaults docDefaults2 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault2 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle2 = new RunPropertiesBaseStyle();
            RunFonts runFonts19 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
            Languages languages30 = new Languages() { Val = "ru-RU", EastAsia = "ru-RU", Bidi = "ar-SA" };

            runPropertiesBaseStyle2.Append(runFonts19);
            runPropertiesBaseStyle2.Append(languages30);

            runPropertiesDefault2.Append(runPropertiesBaseStyle2);
            ParagraphPropertiesDefault paragraphPropertiesDefault2 = new ParagraphPropertiesDefault();

            docDefaults2.Append(runPropertiesDefault2);
            docDefaults2.Append(paragraphPropertiesDefault2);

            LatentStyles latentStyles2 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 99, DefaultSemiHidden = true, DefaultUnhideWhenUsed = true, DefaultPrimaryStyle = false, Count = 267 };
            LatentStyleExceptionInfo latentStyleExceptionInfo139 = new LatentStyleExceptionInfo() { Name = "Normal", UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo140 = new LatentStyleExceptionInfo() { Name = "heading 1", UiPriority = 9, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo141 = new LatentStyleExceptionInfo() { Name = "heading 2", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo142 = new LatentStyleExceptionInfo() { Name = "heading 3", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo143 = new LatentStyleExceptionInfo() { Name = "heading 4", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo144 = new LatentStyleExceptionInfo() { Name = "heading 5", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo145 = new LatentStyleExceptionInfo() { Name = "heading 6", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo146 = new LatentStyleExceptionInfo() { Name = "heading 7", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo147 = new LatentStyleExceptionInfo() { Name = "heading 8", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo148 = new LatentStyleExceptionInfo() { Name = "heading 9", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo149 = new LatentStyleExceptionInfo() { Name = "toc 1", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo150 = new LatentStyleExceptionInfo() { Name = "toc 2", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo151 = new LatentStyleExceptionInfo() { Name = "toc 3", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo152 = new LatentStyleExceptionInfo() { Name = "toc 4", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo153 = new LatentStyleExceptionInfo() { Name = "toc 5", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo154 = new LatentStyleExceptionInfo() { Name = "toc 6", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo155 = new LatentStyleExceptionInfo() { Name = "toc 7", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo156 = new LatentStyleExceptionInfo() { Name = "toc 8", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo157 = new LatentStyleExceptionInfo() { Name = "toc 9", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo158 = new LatentStyleExceptionInfo() { Name = "caption", UiPriority = 35, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo159 = new LatentStyleExceptionInfo() { Name = "Title", UiPriority = 10, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo160 = new LatentStyleExceptionInfo() { Name = "Default Paragraph Font", UiPriority = 1 };
            LatentStyleExceptionInfo latentStyleExceptionInfo161 = new LatentStyleExceptionInfo() { Name = "Subtitle", UiPriority = 11, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo162 = new LatentStyleExceptionInfo() { Name = "Hyperlink", UiPriority = 0 };
            LatentStyleExceptionInfo latentStyleExceptionInfo163 = new LatentStyleExceptionInfo() { Name = "Strong", UiPriority = 22, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo164 = new LatentStyleExceptionInfo() { Name = "Emphasis", UiPriority = 20, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo165 = new LatentStyleExceptionInfo() { Name = "Table Grid", UiPriority = 59, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo166 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo167 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo168 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo169 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo170 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo171 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo172 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo173 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo174 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo175 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo176 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo177 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo178 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo179 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo180 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo181 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo182 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo183 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo184 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo185 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo186 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo187 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo188 = new LatentStyleExceptionInfo() { Name = "Revision", UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo189 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo190 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo191 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo192 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo193 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo194 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo195 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo196 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo197 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo198 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo199 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo200 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo201 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo202 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo203 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo204 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo205 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo206 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo207 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo208 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo209 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo210 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo211 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo212 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo213 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo214 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo215 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo216 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo217 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo218 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo219 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo220 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo221 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo222 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo223 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo224 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo225 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo226 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo227 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo228 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo229 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo230 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo231 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo232 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo233 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo234 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo235 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo236 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo237 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo238 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo239 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo240 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo241 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo242 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo243 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo244 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo245 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo246 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo247 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo248 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo249 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo250 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo251 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo252 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo253 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo254 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo255 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo256 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo257 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo258 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo259 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo260 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo261 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo262 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo263 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo264 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo265 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo266 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo267 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo268 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo269 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo270 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo271 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo272 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo273 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo274 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo275 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37 };
            LatentStyleExceptionInfo latentStyleExceptionInfo276 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, PrimaryStyle = true };

            latentStyles2.Append(latentStyleExceptionInfo139);
            latentStyles2.Append(latentStyleExceptionInfo140);
            latentStyles2.Append(latentStyleExceptionInfo141);
            latentStyles2.Append(latentStyleExceptionInfo142);
            latentStyles2.Append(latentStyleExceptionInfo143);
            latentStyles2.Append(latentStyleExceptionInfo144);
            latentStyles2.Append(latentStyleExceptionInfo145);
            latentStyles2.Append(latentStyleExceptionInfo146);
            latentStyles2.Append(latentStyleExceptionInfo147);
            latentStyles2.Append(latentStyleExceptionInfo148);
            latentStyles2.Append(latentStyleExceptionInfo149);
            latentStyles2.Append(latentStyleExceptionInfo150);
            latentStyles2.Append(latentStyleExceptionInfo151);
            latentStyles2.Append(latentStyleExceptionInfo152);
            latentStyles2.Append(latentStyleExceptionInfo153);
            latentStyles2.Append(latentStyleExceptionInfo154);
            latentStyles2.Append(latentStyleExceptionInfo155);
            latentStyles2.Append(latentStyleExceptionInfo156);
            latentStyles2.Append(latentStyleExceptionInfo157);
            latentStyles2.Append(latentStyleExceptionInfo158);
            latentStyles2.Append(latentStyleExceptionInfo159);
            latentStyles2.Append(latentStyleExceptionInfo160);
            latentStyles2.Append(latentStyleExceptionInfo161);
            latentStyles2.Append(latentStyleExceptionInfo162);
            latentStyles2.Append(latentStyleExceptionInfo163);
            latentStyles2.Append(latentStyleExceptionInfo164);
            latentStyles2.Append(latentStyleExceptionInfo165);
            latentStyles2.Append(latentStyleExceptionInfo166);
            latentStyles2.Append(latentStyleExceptionInfo167);
            latentStyles2.Append(latentStyleExceptionInfo168);
            latentStyles2.Append(latentStyleExceptionInfo169);
            latentStyles2.Append(latentStyleExceptionInfo170);
            latentStyles2.Append(latentStyleExceptionInfo171);
            latentStyles2.Append(latentStyleExceptionInfo172);
            latentStyles2.Append(latentStyleExceptionInfo173);
            latentStyles2.Append(latentStyleExceptionInfo174);
            latentStyles2.Append(latentStyleExceptionInfo175);
            latentStyles2.Append(latentStyleExceptionInfo176);
            latentStyles2.Append(latentStyleExceptionInfo177);
            latentStyles2.Append(latentStyleExceptionInfo178);
            latentStyles2.Append(latentStyleExceptionInfo179);
            latentStyles2.Append(latentStyleExceptionInfo180);
            latentStyles2.Append(latentStyleExceptionInfo181);
            latentStyles2.Append(latentStyleExceptionInfo182);
            latentStyles2.Append(latentStyleExceptionInfo183);
            latentStyles2.Append(latentStyleExceptionInfo184);
            latentStyles2.Append(latentStyleExceptionInfo185);
            latentStyles2.Append(latentStyleExceptionInfo186);
            latentStyles2.Append(latentStyleExceptionInfo187);
            latentStyles2.Append(latentStyleExceptionInfo188);
            latentStyles2.Append(latentStyleExceptionInfo189);
            latentStyles2.Append(latentStyleExceptionInfo190);
            latentStyles2.Append(latentStyleExceptionInfo191);
            latentStyles2.Append(latentStyleExceptionInfo192);
            latentStyles2.Append(latentStyleExceptionInfo193);
            latentStyles2.Append(latentStyleExceptionInfo194);
            latentStyles2.Append(latentStyleExceptionInfo195);
            latentStyles2.Append(latentStyleExceptionInfo196);
            latentStyles2.Append(latentStyleExceptionInfo197);
            latentStyles2.Append(latentStyleExceptionInfo198);
            latentStyles2.Append(latentStyleExceptionInfo199);
            latentStyles2.Append(latentStyleExceptionInfo200);
            latentStyles2.Append(latentStyleExceptionInfo201);
            latentStyles2.Append(latentStyleExceptionInfo202);
            latentStyles2.Append(latentStyleExceptionInfo203);
            latentStyles2.Append(latentStyleExceptionInfo204);
            latentStyles2.Append(latentStyleExceptionInfo205);
            latentStyles2.Append(latentStyleExceptionInfo206);
            latentStyles2.Append(latentStyleExceptionInfo207);
            latentStyles2.Append(latentStyleExceptionInfo208);
            latentStyles2.Append(latentStyleExceptionInfo209);
            latentStyles2.Append(latentStyleExceptionInfo210);
            latentStyles2.Append(latentStyleExceptionInfo211);
            latentStyles2.Append(latentStyleExceptionInfo212);
            latentStyles2.Append(latentStyleExceptionInfo213);
            latentStyles2.Append(latentStyleExceptionInfo214);
            latentStyles2.Append(latentStyleExceptionInfo215);
            latentStyles2.Append(latentStyleExceptionInfo216);
            latentStyles2.Append(latentStyleExceptionInfo217);
            latentStyles2.Append(latentStyleExceptionInfo218);
            latentStyles2.Append(latentStyleExceptionInfo219);
            latentStyles2.Append(latentStyleExceptionInfo220);
            latentStyles2.Append(latentStyleExceptionInfo221);
            latentStyles2.Append(latentStyleExceptionInfo222);
            latentStyles2.Append(latentStyleExceptionInfo223);
            latentStyles2.Append(latentStyleExceptionInfo224);
            latentStyles2.Append(latentStyleExceptionInfo225);
            latentStyles2.Append(latentStyleExceptionInfo226);
            latentStyles2.Append(latentStyleExceptionInfo227);
            latentStyles2.Append(latentStyleExceptionInfo228);
            latentStyles2.Append(latentStyleExceptionInfo229);
            latentStyles2.Append(latentStyleExceptionInfo230);
            latentStyles2.Append(latentStyleExceptionInfo231);
            latentStyles2.Append(latentStyleExceptionInfo232);
            latentStyles2.Append(latentStyleExceptionInfo233);
            latentStyles2.Append(latentStyleExceptionInfo234);
            latentStyles2.Append(latentStyleExceptionInfo235);
            latentStyles2.Append(latentStyleExceptionInfo236);
            latentStyles2.Append(latentStyleExceptionInfo237);
            latentStyles2.Append(latentStyleExceptionInfo238);
            latentStyles2.Append(latentStyleExceptionInfo239);
            latentStyles2.Append(latentStyleExceptionInfo240);
            latentStyles2.Append(latentStyleExceptionInfo241);
            latentStyles2.Append(latentStyleExceptionInfo242);
            latentStyles2.Append(latentStyleExceptionInfo243);
            latentStyles2.Append(latentStyleExceptionInfo244);
            latentStyles2.Append(latentStyleExceptionInfo245);
            latentStyles2.Append(latentStyleExceptionInfo246);
            latentStyles2.Append(latentStyleExceptionInfo247);
            latentStyles2.Append(latentStyleExceptionInfo248);
            latentStyles2.Append(latentStyleExceptionInfo249);
            latentStyles2.Append(latentStyleExceptionInfo250);
            latentStyles2.Append(latentStyleExceptionInfo251);
            latentStyles2.Append(latentStyleExceptionInfo252);
            latentStyles2.Append(latentStyleExceptionInfo253);
            latentStyles2.Append(latentStyleExceptionInfo254);
            latentStyles2.Append(latentStyleExceptionInfo255);
            latentStyles2.Append(latentStyleExceptionInfo256);
            latentStyles2.Append(latentStyleExceptionInfo257);
            latentStyles2.Append(latentStyleExceptionInfo258);
            latentStyles2.Append(latentStyleExceptionInfo259);
            latentStyles2.Append(latentStyleExceptionInfo260);
            latentStyles2.Append(latentStyleExceptionInfo261);
            latentStyles2.Append(latentStyleExceptionInfo262);
            latentStyles2.Append(latentStyleExceptionInfo263);
            latentStyles2.Append(latentStyleExceptionInfo264);
            latentStyles2.Append(latentStyleExceptionInfo265);
            latentStyles2.Append(latentStyleExceptionInfo266);
            latentStyles2.Append(latentStyleExceptionInfo267);
            latentStyles2.Append(latentStyleExceptionInfo268);
            latentStyles2.Append(latentStyleExceptionInfo269);
            latentStyles2.Append(latentStyleExceptionInfo270);
            latentStyles2.Append(latentStyleExceptionInfo271);
            latentStyles2.Append(latentStyleExceptionInfo272);
            latentStyles2.Append(latentStyleExceptionInfo273);
            latentStyles2.Append(latentStyleExceptionInfo274);
            latentStyles2.Append(latentStyleExceptionInfo275);
            latentStyles2.Append(latentStyleExceptionInfo276);

            Style style22 = new Style() { Type = StyleValues.Paragraph, StyleId = "a", Default = true };
            StyleName styleName22 = new StyleName() { Val = "Normal" };
            Rsid rsid19 = new Rsid() { Val = "00831455" };

            StyleRunProperties styleRunProperties17 = new StyleRunProperties();
            RunFonts runFonts20 = new RunFonts() { Ascii = "Arial Unicode MS", HighAnsi = "Arial Unicode MS", EastAsia = "Arial Unicode MS", ComplexScript = "Arial Unicode MS" };
            Color color11 = new Color() { Val = "000000" };
            FontSize fontSize16 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "24" };
            Languages languages31 = new Languages() { Val = "ru" };

            styleRunProperties17.Append(runFonts20);
            styleRunProperties17.Append(color11);
            styleRunProperties17.Append(fontSize16);
            styleRunProperties17.Append(fontSizeComplexScript16);
            styleRunProperties17.Append(languages31);

            style22.Append(styleName22);
            style22.Append(rsid19);
            style22.Append(styleRunProperties17);

            Style style23 = new Style() { Type = StyleValues.Paragraph, StyleId = "1" };
            StyleName styleName23 = new StyleName() { Val = "heading 1" };
            BasedOn basedOn11 = new BasedOn() { Val = "a" };
            NextParagraphStyle nextParagraphStyle2 = new NextParagraphStyle() { Val = "a" };
            LinkedStyle linkedStyle7 = new LinkedStyle() { Val = "10" };
            UIPriority uIPriority11 = new UIPriority() { Val = 9 };
            PrimaryStyle primaryStyle5 = new PrimaryStyle();
            Rsid rsid20 = new Rsid() { Val = "00D37E6B" };

            StyleParagraphProperties styleParagraphProperties7 = new StyleParagraphProperties();
            KeepNext keepNext2 = new KeepNext();
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() { Before = "240", After = "60" };
            OutlineLevel outlineLevel2 = new OutlineLevel() { Val = 0 };

            styleParagraphProperties7.Append(keepNext2);
            styleParagraphProperties7.Append(spacingBetweenLines5);
            styleParagraphProperties7.Append(outlineLevel2);

            StyleRunProperties styleRunProperties18 = new StyleRunProperties();
            RunFonts runFonts21 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold6 = new Bold();
            BoldComplexScript boldComplexScript6 = new BoldComplexScript();
            Kern kern3 = new Kern() { Val = (UInt32Value)32U };
            FontSize fontSize17 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties18.Append(runFonts21);
            styleRunProperties18.Append(bold6);
            styleRunProperties18.Append(boldComplexScript6);
            styleRunProperties18.Append(kern3);
            styleRunProperties18.Append(fontSize17);
            styleRunProperties18.Append(fontSizeComplexScript17);

            style23.Append(styleName23);
            style23.Append(basedOn11);
            style23.Append(nextParagraphStyle2);
            style23.Append(linkedStyle7);
            style23.Append(uIPriority11);
            style23.Append(primaryStyle5);
            style23.Append(rsid20);
            style23.Append(styleParagraphProperties7);
            style23.Append(styleRunProperties18);

            Style style24 = new Style() { Type = StyleValues.Character, StyleId = "a0", Default = true };
            StyleName styleName24 = new StyleName() { Val = "Default Paragraph Font" };
            UIPriority uIPriority12 = new UIPriority() { Val = 1 };
            UnhideWhenUsed unhideWhenUsed5 = new UnhideWhenUsed();

            style24.Append(styleName24);
            style24.Append(uIPriority12);
            style24.Append(unhideWhenUsed5);

            Style style25 = new Style() { Type = StyleValues.Table, StyleId = "a1", Default = true };
            StyleName styleName25 = new StyleName() { Val = "Normal Table" };
            UIPriority uIPriority13 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden4 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed6 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle6 = new PrimaryStyle();

            StyleTableProperties styleTableProperties3 = new StyleTableProperties();
            TableIndentation tableIndentation4 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault6 = new TableCellMarginDefault();
            TopMargin topMargin5 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin3 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin5 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin3 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault6.Append(topMargin5);
            tableCellMarginDefault6.Append(tableCellLeftMargin3);
            tableCellMarginDefault6.Append(bottomMargin5);
            tableCellMarginDefault6.Append(tableCellRightMargin3);

            styleTableProperties3.Append(tableIndentation4);
            styleTableProperties3.Append(tableCellMarginDefault6);

            style25.Append(styleName25);
            style25.Append(uIPriority13);
            style25.Append(semiHidden4);
            style25.Append(unhideWhenUsed6);
            style25.Append(primaryStyle6);
            style25.Append(styleTableProperties3);

            Style style26 = new Style() { Type = StyleValues.Numbering, StyleId = "a2", Default = true };
            StyleName styleName26 = new StyleName() { Val = "No List" };
            UIPriority uIPriority14 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden5 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed7 = new UnhideWhenUsed();

            style26.Append(styleName26);
            style26.Append(uIPriority14);
            style26.Append(semiHidden5);
            style26.Append(unhideWhenUsed7);

            Style style27 = new Style() { Type = StyleValues.Paragraph, StyleId = "a3" };
            StyleName styleName27 = new StyleName() { Val = "No Spacing" };
            UIPriority uIPriority15 = new UIPriority() { Val = 1 };
            PrimaryStyle primaryStyle7 = new PrimaryStyle();
            Rsid rsid21 = new Rsid() { Val = "00D37E6B" };

            StyleRunProperties styleRunProperties19 = new StyleRunProperties();
            RunFonts runFonts22 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
            FontSize fontSize18 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "22" };
            Languages languages32 = new Languages() { EastAsia = "en-US" };

            styleRunProperties19.Append(runFonts22);
            styleRunProperties19.Append(fontSize18);
            styleRunProperties19.Append(fontSizeComplexScript18);
            styleRunProperties19.Append(languages32);

            style27.Append(styleName27);
            style27.Append(uIPriority15);
            style27.Append(primaryStyle7);
            style27.Append(rsid21);
            style27.Append(styleRunProperties19);

            Style style28 = new Style() { Type = StyleValues.Character, StyleId = "a4", CustomStyle = true };
            StyleName styleName28 = new StyleName() { Val = "Сноска_" };
            LinkedStyle linkedStyle8 = new LinkedStyle() { Val = "a5" };
            Rsid rsid22 = new Rsid() { Val = "00831455" };

            StyleRunProperties styleRunProperties20 = new StyleRunProperties();
            RunFonts runFonts23 = new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI", EastAsia = "Segoe UI", ComplexScript = "Segoe UI" };
            FontSize fontSize19 = new FontSize() { Val = "19" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "19" };
            Shading shading23 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            styleRunProperties20.Append(runFonts23);
            styleRunProperties20.Append(fontSize19);
            styleRunProperties20.Append(fontSizeComplexScript19);
            styleRunProperties20.Append(shading23);

            style28.Append(styleName28);
            style28.Append(linkedStyle8);
            style28.Append(rsid22);
            style28.Append(styleRunProperties20);

            Style style29 = new Style() { Type = StyleValues.Character, StyleId = "a6", CustomStyle = true };
            StyleName styleName29 = new StyleName() { Val = "Основной текст_" };
            LinkedStyle linkedStyle9 = new LinkedStyle() { Val = "11" };
            Rsid rsid23 = new Rsid() { Val = "00831455" };

            StyleRunProperties styleRunProperties21 = new StyleRunProperties();
            RunFonts runFonts24 = new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI", EastAsia = "Segoe UI", ComplexScript = "Segoe UI" };
            FontSize fontSize20 = new FontSize() { Val = "19" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "19" };
            Shading shading24 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            styleRunProperties21.Append(runFonts24);
            styleRunProperties21.Append(fontSize20);
            styleRunProperties21.Append(fontSizeComplexScript20);
            styleRunProperties21.Append(shading24);

            style29.Append(styleName29);
            style29.Append(linkedStyle9);
            style29.Append(rsid23);
            style29.Append(styleRunProperties21);

            Style style30 = new Style() { Type = StyleValues.Paragraph, StyleId = "a5", CustomStyle = true };
            StyleName styleName30 = new StyleName() { Val = "Сноска" };
            BasedOn basedOn12 = new BasedOn() { Val = "a" };
            LinkedStyle linkedStyle10 = new LinkedStyle() { Val = "a4" };
            Rsid rsid24 = new Rsid() { Val = "00831455" };

            StyleParagraphProperties styleParagraphProperties8 = new StyleParagraphProperties();
            Shading shading25 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines() { Line = "0", LineRule = LineSpacingRuleValues.AtLeast };

            styleParagraphProperties8.Append(shading25);
            styleParagraphProperties8.Append(spacingBetweenLines6);

            StyleRunProperties styleRunProperties22 = new StyleRunProperties();
            RunFonts runFonts25 = new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI", EastAsia = "Segoe UI", ComplexScript = "Times New Roman" };
            Color color12 = new Color() { Val = "auto" };
            FontSize fontSize21 = new FontSize() { Val = "19" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "19" };
            Languages languages33 = new Languages() { Val = "x-none", EastAsia = "x-none" };

            styleRunProperties22.Append(runFonts25);
            styleRunProperties22.Append(color12);
            styleRunProperties22.Append(fontSize21);
            styleRunProperties22.Append(fontSizeComplexScript21);
            styleRunProperties22.Append(languages33);

            style30.Append(styleName30);
            style30.Append(basedOn12);
            style30.Append(linkedStyle10);
            style30.Append(rsid24);
            style30.Append(styleParagraphProperties8);
            style30.Append(styleRunProperties22);

            Style style31 = new Style() { Type = StyleValues.Paragraph, StyleId = "11", CustomStyle = true };
            StyleName styleName31 = new StyleName() { Val = "Основной текст1" };
            BasedOn basedOn13 = new BasedOn() { Val = "a" };
            LinkedStyle linkedStyle11 = new LinkedStyle() { Val = "a6" };
            Rsid rsid25 = new Rsid() { Val = "00831455" };

            StyleParagraphProperties styleParagraphProperties9 = new StyleParagraphProperties();
            Shading shading26 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines() { Line = "0", LineRule = LineSpacingRuleValues.AtLeast };

            styleParagraphProperties9.Append(shading26);
            styleParagraphProperties9.Append(spacingBetweenLines7);

            StyleRunProperties styleRunProperties23 = new StyleRunProperties();
            RunFonts runFonts26 = new RunFonts() { Ascii = "Segoe UI", HighAnsi = "Segoe UI", EastAsia = "Segoe UI", ComplexScript = "Times New Roman" };
            Color color13 = new Color() { Val = "auto" };
            FontSize fontSize22 = new FontSize() { Val = "19" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "19" };
            Languages languages34 = new Languages() { Val = "x-none", EastAsia = "x-none" };

            styleRunProperties23.Append(runFonts26);
            styleRunProperties23.Append(color13);
            styleRunProperties23.Append(fontSize22);
            styleRunProperties23.Append(fontSizeComplexScript22);
            styleRunProperties23.Append(languages34);

            style31.Append(styleName31);
            style31.Append(basedOn13);
            style31.Append(linkedStyle11);
            style31.Append(rsid25);
            style31.Append(styleParagraphProperties9);
            style31.Append(styleRunProperties23);

            Style style32 = new Style() { Type = StyleValues.Paragraph, StyleId = "2", CustomStyle = true };
            StyleName styleName32 = new StyleName() { Val = "Знак2" };
            BasedOn basedOn14 = new BasedOn() { Val = "a" };
            Rsid rsid26 = new Rsid() { Val = "00831455" };

            StyleParagraphProperties styleParagraphProperties10 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines() { After = "160", Line = "240", LineRule = LineSpacingRuleValues.Exact };

            styleParagraphProperties10.Append(spacingBetweenLines8);

            StyleRunProperties styleRunProperties24 = new StyleRunProperties();
            RunFonts runFonts27 = new RunFonts() { Ascii = "Verdana", HighAnsi = "Verdana", EastAsia = "Times New Roman", ComplexScript = "Verdana" };
            Color color14 = new Color() { Val = "auto" };
            FontSize fontSize23 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "20" };
            Languages languages35 = new Languages() { Val = "en-US", EastAsia = "en-US" };

            styleRunProperties24.Append(runFonts27);
            styleRunProperties24.Append(color14);
            styleRunProperties24.Append(fontSize23);
            styleRunProperties24.Append(fontSizeComplexScript23);
            styleRunProperties24.Append(languages35);

            style32.Append(styleName32);
            style32.Append(basedOn14);
            style32.Append(rsid26);
            style32.Append(styleParagraphProperties10);
            style32.Append(styleRunProperties24);

            Style style33 = new Style() { Type = StyleValues.Character, StyleId = "a7" };
            StyleName styleName33 = new StyleName() { Val = "Hyperlink" };
            Rsid rsid27 = new Rsid() { Val = "00831455" };

            StyleRunProperties styleRunProperties25 = new StyleRunProperties();
            Color color15 = new Color() { Val = "0000FF" };
            Underline underline3 = new Underline() { Val = UnderlineValues.Single };

            styleRunProperties25.Append(color15);
            styleRunProperties25.Append(underline3);

            style33.Append(styleName33);
            style33.Append(rsid27);
            style33.Append(styleRunProperties25);

            Style style34 = new Style() { Type = StyleValues.Paragraph, StyleId = "a8" };
            StyleName styleName34 = new StyleName() { Val = "List Paragraph" };
            BasedOn basedOn15 = new BasedOn() { Val = "a" };
            UIPriority uIPriority16 = new UIPriority() { Val = 34 };
            PrimaryStyle primaryStyle8 = new PrimaryStyle();
            Rsid rsid28 = new Rsid() { Val = "00FE63E5" };

            StyleParagraphProperties styleParagraphProperties11 = new StyleParagraphProperties();
            Indentation indentation2 = new Indentation() { Left = "720" };
            ContextualSpacing contextualSpacing2 = new ContextualSpacing();

            styleParagraphProperties11.Append(indentation2);
            styleParagraphProperties11.Append(contextualSpacing2);

            style34.Append(styleName34);
            style34.Append(basedOn15);
            style34.Append(uIPriority16);
            style34.Append(primaryStyle8);
            style34.Append(rsid28);
            style34.Append(styleParagraphProperties11);

            Style style35 = new Style() { Type = StyleValues.Character, StyleId = "FontStyle77", CustomStyle = true };
            StyleName styleName35 = new StyleName() { Val = "Font Style77" };
            UIPriority uIPriority17 = new UIPriority() { Val = 99 };
            Rsid rsid29 = new Rsid() { Val = "00C52219" };

            StyleRunProperties styleRunProperties26 = new StyleRunProperties();
            RunFonts runFonts28 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold7 = new Bold();
            BoldComplexScript boldComplexScript7 = new BoldComplexScript();
            FontSize fontSize24 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties26.Append(runFonts28);
            styleRunProperties26.Append(bold7);
            styleRunProperties26.Append(boldComplexScript7);
            styleRunProperties26.Append(fontSize24);
            styleRunProperties26.Append(fontSizeComplexScript24);

            style35.Append(styleName35);
            style35.Append(uIPriority17);
            style35.Append(rsid29);
            style35.Append(styleRunProperties26);

            Style style36 = new Style() { Type = StyleValues.Table, StyleId = "a9" };
            StyleName styleName36 = new StyleName() { Val = "Table Grid" };
            BasedOn basedOn16 = new BasedOn() { Val = "a1" };
            UIPriority uIPriority18 = new UIPriority() { Val = 59 };
            Rsid rsid30 = new Rsid() { Val = "00F444F2" };

            StyleTableProperties styleTableProperties4 = new StyleTableProperties();

            TableBorders tableBorders3 = new TableBorders();
            TopBorder topBorder21 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder21 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder21 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder21 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder3 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder3 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders3.Append(topBorder21);
            tableBorders3.Append(leftBorder21);
            tableBorders3.Append(bottomBorder21);
            tableBorders3.Append(rightBorder21);
            tableBorders3.Append(insideHorizontalBorder3);
            tableBorders3.Append(insideVerticalBorder3);

            styleTableProperties4.Append(tableBorders3);

            style36.Append(styleName36);
            style36.Append(basedOn16);
            style36.Append(uIPriority18);
            style36.Append(rsid30);
            style36.Append(styleTableProperties4);

            Style style37 = new Style() { Type = StyleValues.Character, StyleId = "rvts8", CustomStyle = true };
            StyleName styleName37 = new StyleName() { Val = "rvts8" };
            BasedOn basedOn17 = new BasedOn() { Val = "a0" };
            Rsid rsid31 = new Rsid() { Val = "00E86ABA" };

            StyleRunProperties styleRunProperties27 = new StyleRunProperties();
            Bold bold8 = new Bold();
            BoldComplexScript boldComplexScript8 = new BoldComplexScript();
            FontSize fontSize25 = new FontSize() { Val = "36" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "36" };

            styleRunProperties27.Append(bold8);
            styleRunProperties27.Append(boldComplexScript8);
            styleRunProperties27.Append(fontSize25);
            styleRunProperties27.Append(fontSizeComplexScript25);

            style37.Append(styleName37);
            style37.Append(basedOn17);
            style37.Append(rsid31);
            style37.Append(styleRunProperties27);

            Style style38 = new Style() { Type = StyleValues.Paragraph, StyleId = "aa" };
            StyleName styleName38 = new StyleName() { Val = "Normal (Web)" };
            BasedOn basedOn18 = new BasedOn() { Val = "a" };
            UIPriority uIPriority19 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden6 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed8 = new UnhideWhenUsed();
            Rsid rsid32 = new Rsid() { Val = "00E86ABA" };

            StyleRunProperties styleRunProperties28 = new StyleRunProperties();
            RunFonts runFonts29 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Calibri", ComplexScript = "Times New Roman" };
            Color color16 = new Color() { Val = "auto" };
            Languages languages36 = new Languages() { Val = "ru-RU" };

            styleRunProperties28.Append(runFonts29);
            styleRunProperties28.Append(color16);
            styleRunProperties28.Append(languages36);

            style38.Append(styleName38);
            style38.Append(basedOn18);
            style38.Append(uIPriority19);
            style38.Append(semiHidden6);
            style38.Append(unhideWhenUsed8);
            style38.Append(rsid32);
            style38.Append(styleRunProperties28);

            Style style39 = new Style() { Type = StyleValues.Character, StyleId = "rvts6", CustomStyle = true };
            StyleName styleName39 = new StyleName() { Val = "rvts6" };
            BasedOn basedOn19 = new BasedOn() { Val = "a0" };
            Rsid rsid33 = new Rsid() { Val = "00E86ABA" };

            StyleRunProperties styleRunProperties29 = new StyleRunProperties();
            FontSize fontSize26 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "28" };

            styleRunProperties29.Append(fontSize26);
            styleRunProperties29.Append(fontSizeComplexScript26);

            style39.Append(styleName39);
            style39.Append(basedOn19);
            style39.Append(rsid33);
            style39.Append(styleRunProperties29);

            Style style40 = new Style() { Type = StyleValues.Character, StyleId = "12", CustomStyle = true };
            StyleName styleName40 = new StyleName() { Val = "Заголовок №1" };
            Rsid rsid34 = new Rsid() { Val = "00A16BF1" };

            StyleRunProperties styleRunProperties30 = new StyleRunProperties();
            RunFonts runFonts30 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman", ComplexScript = "Times New Roman" };
            Bold bold9 = new Bold();
            BoldComplexScript boldComplexScript9 = new BoldComplexScript();
            Italic italic2 = new Italic() { Val = false };
            ItalicComplexScript italicComplexScript2 = new ItalicComplexScript() { Val = false };
            Caps caps2 = new Caps() { Val = false };
            SmallCaps smallCaps2 = new SmallCaps() { Val = false };
            Strike strike2 = new Strike() { Val = false };
            DoubleStrike doubleStrike2 = new DoubleStrike() { Val = false };
            Color color17 = new Color() { Val = "000000" };
            Spacing spacing2 = new Spacing() { Val = 18 };
            CharacterScale characterScale2 = new CharacterScale() { Val = 100 };
            Position position2 = new Position() { Val = "0" };
            FontSize fontSize27 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "23" };
            Underline underline4 = new Underline() { Val = UnderlineValues.Single };
            VerticalTextAlignment verticalTextAlignment2 = new VerticalTextAlignment() { Val = VerticalPositionValues.Baseline };
            Languages languages37 = new Languages() { Val = "ru-RU" };

            styleRunProperties30.Append(runFonts30);
            styleRunProperties30.Append(bold9);
            styleRunProperties30.Append(boldComplexScript9);
            styleRunProperties30.Append(italic2);
            styleRunProperties30.Append(italicComplexScript2);
            styleRunProperties30.Append(caps2);
            styleRunProperties30.Append(smallCaps2);
            styleRunProperties30.Append(strike2);
            styleRunProperties30.Append(doubleStrike2);
            styleRunProperties30.Append(color17);
            styleRunProperties30.Append(spacing2);
            styleRunProperties30.Append(characterScale2);
            styleRunProperties30.Append(position2);
            styleRunProperties30.Append(fontSize27);
            styleRunProperties30.Append(fontSizeComplexScript27);
            styleRunProperties30.Append(underline4);
            styleRunProperties30.Append(verticalTextAlignment2);
            styleRunProperties30.Append(languages37);

            style40.Append(styleName40);
            style40.Append(rsid34);
            style40.Append(styleRunProperties30);

            Style style41 = new Style() { Type = StyleValues.Paragraph, StyleId = "Default", CustomStyle = true };
            StyleName styleName41 = new StyleName() { Val = "Default" };
            Rsid rsid35 = new Rsid() { Val = "002D5ECD" };

            StyleParagraphProperties styleParagraphProperties12 = new StyleParagraphProperties();
            AutoSpaceDE autoSpaceDE2 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN2 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent2 = new AdjustRightIndent() { Val = false };

            styleParagraphProperties12.Append(autoSpaceDE2);
            styleParagraphProperties12.Append(autoSpaceDN2);
            styleParagraphProperties12.Append(adjustRightIndent2);

            StyleRunProperties styleRunProperties31 = new StyleRunProperties();
            RunFonts runFonts31 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "Times New Roman" };
            Color color18 = new Color() { Val = "000000" };
            FontSize fontSize28 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties31.Append(runFonts31);
            styleRunProperties31.Append(color18);
            styleRunProperties31.Append(fontSize28);
            styleRunProperties31.Append(fontSizeComplexScript28);

            style41.Append(styleName41);
            style41.Append(rsid35);
            style41.Append(styleParagraphProperties12);
            style41.Append(styleRunProperties31);

            Style style42 = new Style() { Type = StyleValues.Character, StyleId = "10", CustomStyle = true };
            StyleName styleName42 = new StyleName() { Val = "Заголовок 1 Знак" };
            BasedOn basedOn20 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle12 = new LinkedStyle() { Val = "1" };
            UIPriority uIPriority20 = new UIPriority() { Val = 9 };
            Rsid rsid36 = new Rsid() { Val = "00D37E6B" };

            StyleRunProperties styleRunProperties32 = new StyleRunProperties();
            RunFonts runFonts32 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold10 = new Bold();
            BoldComplexScript boldComplexScript10 = new BoldComplexScript();
            Color color19 = new Color() { Val = "000000" };
            Kern kern4 = new Kern() { Val = (UInt32Value)32U };
            FontSize fontSize29 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "32" };
            Languages languages38 = new Languages() { Val = "ru" };

            styleRunProperties32.Append(runFonts32);
            styleRunProperties32.Append(bold10);
            styleRunProperties32.Append(boldComplexScript10);
            styleRunProperties32.Append(color19);
            styleRunProperties32.Append(kern4);
            styleRunProperties32.Append(fontSize29);
            styleRunProperties32.Append(fontSizeComplexScript29);
            styleRunProperties32.Append(languages38);

            style42.Append(styleName42);
            style42.Append(basedOn20);
            style42.Append(linkedStyle12);
            style42.Append(uIPriority20);
            style42.Append(rsid36);
            style42.Append(styleRunProperties32);

            styles2.Append(docDefaults2);
            styles2.Append(latentStyles2);
            styles2.Append(style22);
            styles2.Append(style23);
            styles2.Append(style24);
            styles2.Append(style25);
            styles2.Append(style26);
            styles2.Append(style27);
            styles2.Append(style28);
            styles2.Append(style29);
            styles2.Append(style30);
            styles2.Append(style31);
            styles2.Append(style32);
            styles2.Append(style33);
            styles2.Append(style34);
            styles2.Append(style35);
            styles2.Append(style36);
            styles2.Append(style37);
            styles2.Append(style38);
            styles2.Append(style39);
            styles2.Append(style40);
            styles2.Append(style41);
            styles2.Append(style42);

            styleDefinitionsPart1.Styles = styles2;
        }

        // Generates content of numberingDefinitionsPart1.
        private void GenerateNumberingDefinitionsPart1Content(NumberingDefinitionsPart numberingDefinitionsPart1)
        {
            Numbering numbering1 = new Numbering() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            numbering1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            numbering1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            numbering1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            numbering1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            numbering1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            numbering1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            numbering1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            numbering1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            numbering1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            numbering1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            numbering1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            numbering1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            numbering1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            numbering1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            numbering1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = 0 };
            Nsid nsid1 = new Nsid() { Val = "5C7F2C53" };
            MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode1 = new TemplateCode() { Val = "607E5354" };

            Level level1 = new Level() { LevelIndex = 0, TemplateCode = "0419000F" };
            StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText1 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
            Indentation indentation3 = new Indentation() { Left = "720", Hanging = "360" };

            previousParagraphProperties1.Append(indentation3);

            level1.Append(startNumberingValue1);
            level1.Append(numberingFormat1);
            level1.Append(levelText1);
            level1.Append(levelJustification1);
            level1.Append(previousParagraphProperties1);

            Level level2 = new Level() { LevelIndex = 1, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText2 = new LevelText() { Val = "%2." };
            LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
            Indentation indentation4 = new Indentation() { Left = "1440", Hanging = "360" };

            previousParagraphProperties2.Append(indentation4);

            level2.Append(startNumberingValue2);
            level2.Append(numberingFormat2);
            level2.Append(levelText2);
            level2.Append(levelJustification2);
            level2.Append(previousParagraphProperties2);

            Level level3 = new Level() { LevelIndex = 2, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText3 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
            Indentation indentation5 = new Indentation() { Left = "2160", Hanging = "180" };

            previousParagraphProperties3.Append(indentation5);

            level3.Append(startNumberingValue3);
            level3.Append(numberingFormat3);
            level3.Append(levelText3);
            level3.Append(levelJustification3);
            level3.Append(previousParagraphProperties3);

            Level level4 = new Level() { LevelIndex = 3, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText4 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
            Indentation indentation6 = new Indentation() { Left = "2880", Hanging = "360" };

            previousParagraphProperties4.Append(indentation6);

            level4.Append(startNumberingValue4);
            level4.Append(numberingFormat4);
            level4.Append(levelText4);
            level4.Append(levelJustification4);
            level4.Append(previousParagraphProperties4);

            Level level5 = new Level() { LevelIndex = 4, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText5 = new LevelText() { Val = "%5." };
            LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
            Indentation indentation7 = new Indentation() { Left = "3600", Hanging = "360" };

            previousParagraphProperties5.Append(indentation7);

            level5.Append(startNumberingValue5);
            level5.Append(numberingFormat5);
            level5.Append(levelText5);
            level5.Append(levelJustification5);
            level5.Append(previousParagraphProperties5);

            Level level6 = new Level() { LevelIndex = 5, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText6 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
            Indentation indentation8 = new Indentation() { Left = "4320", Hanging = "180" };

            previousParagraphProperties6.Append(indentation8);

            level6.Append(startNumberingValue6);
            level6.Append(numberingFormat6);
            level6.Append(levelText6);
            level6.Append(levelJustification6);
            level6.Append(previousParagraphProperties6);

            Level level7 = new Level() { LevelIndex = 6, TemplateCode = "0419000F", Tentative = true };
            StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText7 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
            Indentation indentation9 = new Indentation() { Left = "5040", Hanging = "360" };

            previousParagraphProperties7.Append(indentation9);

            level7.Append(startNumberingValue7);
            level7.Append(numberingFormat7);
            level7.Append(levelText7);
            level7.Append(levelJustification7);
            level7.Append(previousParagraphProperties7);

            Level level8 = new Level() { LevelIndex = 7, TemplateCode = "04190019", Tentative = true };
            StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
            LevelText levelText8 = new LevelText() { Val = "%8." };
            LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
            Indentation indentation10 = new Indentation() { Left = "5760", Hanging = "360" };

            previousParagraphProperties8.Append(indentation10);

            level8.Append(startNumberingValue8);
            level8.Append(numberingFormat8);
            level8.Append(levelText8);
            level8.Append(levelJustification8);
            level8.Append(previousParagraphProperties8);

            Level level9 = new Level() { LevelIndex = 8, TemplateCode = "0419001B", Tentative = true };
            StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText9 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
            Indentation indentation11 = new Indentation() { Left = "6480", Hanging = "180" };

            previousParagraphProperties9.Append(indentation11);

            level9.Append(startNumberingValue9);
            level9.Append(numberingFormat9);
            level9.Append(levelText9);
            level9.Append(levelJustification9);
            level9.Append(previousParagraphProperties9);

            abstractNum1.Append(nsid1);
            abstractNum1.Append(multiLevelType1);
            abstractNum1.Append(templateCode1);
            abstractNum1.Append(level1);
            abstractNum1.Append(level2);
            abstractNum1.Append(level3);
            abstractNum1.Append(level4);
            abstractNum1.Append(level5);
            abstractNum1.Append(level6);
            abstractNum1.Append(level7);
            abstractNum1.Append(level8);
            abstractNum1.Append(level9);

            NumberingInstance numberingInstance1 = new NumberingInstance() { NumberID = 1 };
            AbstractNumId abstractNumId1 = new AbstractNumId() { Val = 0 };

            numberingInstance1.Append(abstractNumId1);

            numbering1.Append(abstractNum1);
            numbering1.Append(numberingInstance1);

            numberingDefinitionsPart1.Numbering = numbering1;
        }

        // Generates content of footnotesPart1.
        private void GenerateFootnotesPart1Content(FootnotesPart footnotesPart1)
        {
            Footnotes footnotes1 = new Footnotes() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            footnotes1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            footnotes1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footnotes1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footnotes1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footnotes1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footnotes1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footnotes1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            footnotes1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footnotes1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footnotes1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footnotes1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            footnotes1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            footnotes1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            footnotes1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            footnotes1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Footnote footnote1 = new Footnote() { Type = FootnoteEndnoteValues.Separator, Id = -1 };

            Paragraph paragraph42 = new Paragraph() { RsidParagraphAddition = "00EF038F", RsidParagraphProperties = "00831455", RsidRunAdditionDefault = "00EF038F" };

            Run run98 = new Run();
            SeparatorMark separatorMark2 = new SeparatorMark();

            run98.Append(separatorMark2);

            paragraph42.Append(run98);

            footnote1.Append(paragraph42);

            Footnote footnote2 = new Footnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph43 = new Paragraph() { RsidParagraphAddition = "00EF038F", RsidParagraphProperties = "00831455", RsidRunAdditionDefault = "00EF038F" };

            Run run99 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark2 = new ContinuationSeparatorMark();

            run99.Append(continuationSeparatorMark2);

            paragraph43.Append(run99);

            footnote2.Append(paragraph43);

            footnotes1.Append(footnote1);
            footnotes1.Append(footnote2);

            footnotesPart1.Footnotes = footnotes1;
        }

        // Generates content of webSettingsPart1.
        private void GenerateWebSettingsPart1Content(WebSettingsPart webSettingsPart1)
        {
            WebSettings webSettings1 = new WebSettings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            webSettings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            webSettings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            webSettings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            webSettings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");

            Divs divs1 = new Divs();

            Div div1 = new Div() { Id = "1625035778" };
            BodyDiv bodyDiv1 = new BodyDiv() { Val = true };
            LeftMarginDiv leftMarginDiv1 = new LeftMarginDiv() { Val = "0" };
            RightMarginDiv rightMarginDiv1 = new RightMarginDiv() { Val = "0" };
            TopMarginDiv topMarginDiv1 = new TopMarginDiv() { Val = "0" };
            BottomMarginDiv bottomMarginDiv1 = new BottomMarginDiv() { Val = "0" };

            DivBorder divBorder1 = new DivBorder();
            TopBorder topBorder22 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder22 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder22 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder22 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder1.Append(topBorder22);
            divBorder1.Append(leftBorder22);
            divBorder1.Append(bottomBorder22);
            divBorder1.Append(rightBorder22);

            div1.Append(bodyDiv1);
            div1.Append(leftMarginDiv1);
            div1.Append(rightMarginDiv1);
            div1.Append(topMarginDiv1);
            div1.Append(bottomMarginDiv1);
            div1.Append(divBorder1);

            divs1.Append(div1);
            OptimizeForBrowser optimizeForBrowser1 = new OptimizeForBrowser();
            TargetScreenSize targetScreenSize1 = new TargetScreenSize() { Val = TargetScreenSizeValues.Sz800x600 };

            webSettings1.Append(divs1);
            webSettings1.Append(optimizeForBrowser1);
            webSettings1.Append(targetScreenSize1);

            webSettingsPart1.WebSettings = webSettings1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Тема Office" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Стандартная" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink2 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0000FF" };

            hyperlink2.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "800080" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink2);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Стандартная" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Cambria" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ ゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ 明朝" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Стандартная" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint1 = new A.Tint() { Val = 50000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 35000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint2 = new A.Tint() { Val = 37000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint3 = new A.Tint() { Val = 15000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 16200000, Scaled = true };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade1 = new A.Shade() { Val = 51000 };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 80000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade2 = new A.Shade() { Val = 93000 };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade3 = new A.Shade() { Val = 94000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 16200000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade4 = new A.Shade() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);

            A.Outline outline2 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);

            A.Outline outline3 = new A.Outline() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha2 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha3 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            A.Scene3DType scene3DType1 = new A.Scene3DType();

            A.Camera camera1 = new A.Camera() { Preset = A.PresetCameraValues.OrthographicFront };
            A.Rotation rotation1 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            A.LightRig lightRig1 = new A.LightRig() { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
            A.Rotation rotation2 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            A.Shape3DType shape3DType1 = new A.Shape3DType();
            A.BevelTop bevelTop1 = new A.BevelTop() { Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint4 = new A.Tint() { Val = 40000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop7.Append(schemeColor12);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 40000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 45000 };
            A.Shade shade5 = new A.Shade() { Val = 99000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop8.Append(schemeColor13);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade() { Val = 20000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop9.Append(schemeColor14);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 80000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop10.Append(schemeColor15);

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade7 = new A.Shade() { Val = 30000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop11.Append(schemeColor16);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(gradientFill3);
            backgroundFillStyleList1.Append(gradientFill4);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }

        // Generates content of documentSettingsPart1.
        private void GenerateDocumentSettingsPart1Content(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = new Settings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            settings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            settings1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            settings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            settings1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            settings1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            settings1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            settings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            settings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            settings1.AddNamespaceDeclaration("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
            Zoom zoom1 = new Zoom() { Percent = "100" };
            ProofState proofState1 = new ProofState() { Spelling = ProofingStateValues.Clean, Grammar = ProofingStateValues.Clean };
            DefaultTabStop defaultTabStop1 = new DefaultTabStop() { Val = 737 };
            CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl() { Val = CharacterSpacingValues.DoNotCompress };

            FootnoteDocumentWideProperties footnoteDocumentWideProperties1 = new FootnoteDocumentWideProperties();
            FootnoteSpecialReference footnoteSpecialReference1 = new FootnoteSpecialReference() { Id = -1 };
            FootnoteSpecialReference footnoteSpecialReference2 = new FootnoteSpecialReference() { Id = 0 };

            footnoteDocumentWideProperties1.Append(footnoteSpecialReference1);
            footnoteDocumentWideProperties1.Append(footnoteSpecialReference2);

            EndnoteDocumentWideProperties endnoteDocumentWideProperties1 = new EndnoteDocumentWideProperties();
            EndnoteSpecialReference endnoteSpecialReference1 = new EndnoteSpecialReference() { Id = -1 };
            EndnoteSpecialReference endnoteSpecialReference2 = new EndnoteSpecialReference() { Id = 0 };

            endnoteDocumentWideProperties1.Append(endnoteSpecialReference1);
            endnoteDocumentWideProperties1.Append(endnoteSpecialReference2);

            Compatibility compatibility1 = new Compatibility();
            CompatibilitySetting compatibilitySetting1 = new CompatibilitySetting() { Name = CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word", Val = "14" };
            CompatibilitySetting compatibilitySetting2 = new CompatibilitySetting() { Name = CompatSettingNameValues.OverrideTableStyleFontSizeAndJustification, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting3 = new CompatibilitySetting() { Name = CompatSettingNameValues.EnableOpenTypeFeatures, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting4 = new CompatibilitySetting() { Name = CompatSettingNameValues.DoNotFlipMirrorIndents, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };

            compatibility1.Append(compatibilitySetting1);
            compatibility1.Append(compatibilitySetting2);
            compatibility1.Append(compatibilitySetting3);
            compatibility1.Append(compatibilitySetting4);

            Rsids rsids1 = new Rsids();
            RsidRoot rsidRoot1 = new RsidRoot() { Val = "00831455" };
            Rsid rsid37 = new Rsid() { Val = "00023BBE" };
            Rsid rsid38 = new Rsid() { Val = "00037588" };
            Rsid rsid39 = new Rsid() { Val = "00087587" };
            Rsid rsid40 = new Rsid() { Val = "000E0BE7" };
            Rsid rsid41 = new Rsid() { Val = "0012341B" };
            Rsid rsid42 = new Rsid() { Val = "00143A64" };
            Rsid rsid43 = new Rsid() { Val = "001555D6" };
            Rsid rsid44 = new Rsid() { Val = "001D5FD5" };
            Rsid rsid45 = new Rsid() { Val = "001E2288" };
            Rsid rsid46 = new Rsid() { Val = "001F77F0" };
            Rsid rsid47 = new Rsid() { Val = "0027499D" };
            Rsid rsid48 = new Rsid() { Val = "002811AA" };
            Rsid rsid49 = new Rsid() { Val = "002963B0" };
            Rsid rsid50 = new Rsid() { Val = "002A59F2" };
            Rsid rsid51 = new Rsid() { Val = "002B166B" };
            Rsid rsid52 = new Rsid() { Val = "002B5A19" };
            Rsid rsid53 = new Rsid() { Val = "002D428F" };
            Rsid rsid54 = new Rsid() { Val = "002D5ECD" };
            Rsid rsid55 = new Rsid() { Val = "002E7729" };
            Rsid rsid56 = new Rsid() { Val = "002F6285" };
            Rsid rsid57 = new Rsid() { Val = "00345B20" };
            Rsid rsid58 = new Rsid() { Val = "00384CA5" };
            Rsid rsid59 = new Rsid() { Val = "003B3DC8" };
            Rsid rsid60 = new Rsid() { Val = "005548B6" };
            Rsid rsid61 = new Rsid() { Val = "005E4241" };
            Rsid rsid62 = new Rsid() { Val = "005F5365" };
            Rsid rsid63 = new Rsid() { Val = "00677B86" };
            Rsid rsid64 = new Rsid() { Val = "006E1D99" };
            Rsid rsid65 = new Rsid() { Val = "007F46F0" };
            Rsid rsid66 = new Rsid() { Val = "007F6711" };
            Rsid rsid67 = new Rsid() { Val = "00831455" };
            Rsid rsid68 = new Rsid() { Val = "008935D9" };
            Rsid rsid69 = new Rsid() { Val = "008D46DE" };
            Rsid rsid70 = new Rsid() { Val = "008F5AB8" };
            Rsid rsid71 = new Rsid() { Val = "00926D26" };
            Rsid rsid72 = new Rsid() { Val = "00934C50" };
            Rsid rsid73 = new Rsid() { Val = "00974E84" };
            Rsid rsid74 = new Rsid() { Val = "00993140" };
            Rsid rsid75 = new Rsid() { Val = "009F347A" };
            Rsid rsid76 = new Rsid() { Val = "009F70EB" };
            Rsid rsid77 = new Rsid() { Val = "00A16BF1" };
            Rsid rsid78 = new Rsid() { Val = "00A61CB9" };
            Rsid rsid79 = new Rsid() { Val = "00A72F6F" };
            Rsid rsid80 = new Rsid() { Val = "00AC0B5A" };
            Rsid rsid81 = new Rsid() { Val = "00AD59C7" };
            Rsid rsid82 = new Rsid() { Val = "00B015F4" };
            Rsid rsid83 = new Rsid() { Val = "00B131D1" };
            Rsid rsid84 = new Rsid() { Val = "00B230F0" };
            Rsid rsid85 = new Rsid() { Val = "00B837E2" };
            Rsid rsid86 = new Rsid() { Val = "00C52219" };
            Rsid rsid87 = new Rsid() { Val = "00C72522" };
            Rsid rsid88 = new Rsid() { Val = "00CA334E" };
            Rsid rsid89 = new Rsid() { Val = "00CC6D34" };
            Rsid rsid90 = new Rsid() { Val = "00CE243D" };
            Rsid rsid91 = new Rsid() { Val = "00D37E6B" };
            Rsid rsid92 = new Rsid() { Val = "00DC2B6D" };
            Rsid rsid93 = new Rsid() { Val = "00E139D6" };
            Rsid rsid94 = new Rsid() { Val = "00E46DD7" };
            Rsid rsid95 = new Rsid() { Val = "00E86ABA" };
            Rsid rsid96 = new Rsid() { Val = "00EA270F" };
            Rsid rsid97 = new Rsid() { Val = "00EE305B" };
            Rsid rsid98 = new Rsid() { Val = "00EF038F" };
            Rsid rsid99 = new Rsid() { Val = "00F15D05" };
            Rsid rsid100 = new Rsid() { Val = "00F4161B" };
            Rsid rsid101 = new Rsid() { Val = "00F444F2" };
            Rsid rsid102 = new Rsid() { Val = "00F615DE" };
            Rsid rsid103 = new Rsid() { Val = "00FE63E5" };

            rsids1.Append(rsidRoot1);
            rsids1.Append(rsid37);
            rsids1.Append(rsid38);
            rsids1.Append(rsid39);
            rsids1.Append(rsid40);
            rsids1.Append(rsid41);
            rsids1.Append(rsid42);
            rsids1.Append(rsid43);
            rsids1.Append(rsid44);
            rsids1.Append(rsid45);
            rsids1.Append(rsid46);
            rsids1.Append(rsid47);
            rsids1.Append(rsid48);
            rsids1.Append(rsid49);
            rsids1.Append(rsid50);
            rsids1.Append(rsid51);
            rsids1.Append(rsid52);
            rsids1.Append(rsid53);
            rsids1.Append(rsid54);
            rsids1.Append(rsid55);
            rsids1.Append(rsid56);
            rsids1.Append(rsid57);
            rsids1.Append(rsid58);
            rsids1.Append(rsid59);
            rsids1.Append(rsid60);
            rsids1.Append(rsid61);
            rsids1.Append(rsid62);
            rsids1.Append(rsid63);
            rsids1.Append(rsid64);
            rsids1.Append(rsid65);
            rsids1.Append(rsid66);
            rsids1.Append(rsid67);
            rsids1.Append(rsid68);
            rsids1.Append(rsid69);
            rsids1.Append(rsid70);
            rsids1.Append(rsid71);
            rsids1.Append(rsid72);
            rsids1.Append(rsid73);
            rsids1.Append(rsid74);
            rsids1.Append(rsid75);
            rsids1.Append(rsid76);
            rsids1.Append(rsid77);
            rsids1.Append(rsid78);
            rsids1.Append(rsid79);
            rsids1.Append(rsid80);
            rsids1.Append(rsid81);
            rsids1.Append(rsid82);
            rsids1.Append(rsid83);
            rsids1.Append(rsid84);
            rsids1.Append(rsid85);
            rsids1.Append(rsid86);
            rsids1.Append(rsid87);
            rsids1.Append(rsid88);
            rsids1.Append(rsid89);
            rsids1.Append(rsid90);
            rsids1.Append(rsid91);
            rsids1.Append(rsid92);
            rsids1.Append(rsid93);
            rsids1.Append(rsid94);
            rsids1.Append(rsid95);
            rsids1.Append(rsid96);
            rsids1.Append(rsid97);
            rsids1.Append(rsid98);
            rsids1.Append(rsid99);
            rsids1.Append(rsid100);
            rsids1.Append(rsid101);
            rsids1.Append(rsid102);
            rsids1.Append(rsid103);

            M.MathProperties mathProperties1 = new M.MathProperties();
            M.MathFont mathFont1 = new M.MathFont() { Val = "Cambria Math" };
            M.BreakBinary breakBinary1 = new M.BreakBinary() { Val = M.BreakBinaryOperatorValues.Before };
            M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction() { Val = M.BreakBinarySubtractionValues.MinusMinus };
            M.SmallFraction smallFraction1 = new M.SmallFraction() { Val = M.BooleanValues.Zero };
            M.DisplayDefaults displayDefaults1 = new M.DisplayDefaults();
            M.LeftMargin leftMargin1 = new M.LeftMargin() { Val = (UInt32Value)0U };
            M.RightMargin rightMargin1 = new M.RightMargin() { Val = (UInt32Value)0U };
            M.DefaultJustification defaultJustification1 = new M.DefaultJustification() { Val = M.JustificationValues.CenterGroup };
            M.WrapIndent wrapIndent1 = new M.WrapIndent() { Val = (UInt32Value)1440U };
            M.IntegralLimitLocation integralLimitLocation1 = new M.IntegralLimitLocation() { Val = M.LimitLocationValues.SubscriptSuperscript };
            M.NaryLimitLocation naryLimitLocation1 = new M.NaryLimitLocation() { Val = M.LimitLocationValues.UnderOver };

            mathProperties1.Append(mathFont1);
            mathProperties1.Append(breakBinary1);
            mathProperties1.Append(breakBinarySubtraction1);
            mathProperties1.Append(smallFraction1);
            mathProperties1.Append(displayDefaults1);
            mathProperties1.Append(leftMargin1);
            mathProperties1.Append(rightMargin1);
            mathProperties1.Append(defaultJustification1);
            mathProperties1.Append(wrapIndent1);
            mathProperties1.Append(integralLimitLocation1);
            mathProperties1.Append(naryLimitLocation1);
            ThemeFontLanguages themeFontLanguages1 = new ThemeFontLanguages() { Val = "ru-RU" };
            ColorSchemeMapping colorSchemeMapping1 = new ColorSchemeMapping() { Background1 = ColorSchemeIndexValues.Light1, Text1 = ColorSchemeIndexValues.Dark1, Background2 = ColorSchemeIndexValues.Light2, Text2 = ColorSchemeIndexValues.Dark2, Accent1 = ColorSchemeIndexValues.Accent1, Accent2 = ColorSchemeIndexValues.Accent2, Accent3 = ColorSchemeIndexValues.Accent3, Accent4 = ColorSchemeIndexValues.Accent4, Accent5 = ColorSchemeIndexValues.Accent5, Accent6 = ColorSchemeIndexValues.Accent6, Hyperlink = ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink };

            ShapeDefaults shapeDefaults1 = new ShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults2 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 1026 };

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "1" };

            shapeLayout1.Append(shapeIdMap1);

            shapeDefaults1.Append(shapeDefaults2);
            shapeDefaults1.Append(shapeLayout1);
            DecimalSymbol decimalSymbol1 = new DecimalSymbol() { Val = "," };
            ListSeparator listSeparator1 = new ListSeparator() { Val = ";" };

            settings1.Append(zoom1);
            settings1.Append(proofState1);
            settings1.Append(defaultTabStop1);
            settings1.Append(characterSpacingControl1);
            settings1.Append(footnoteDocumentWideProperties1);
            settings1.Append(endnoteDocumentWideProperties1);
            settings1.Append(compatibility1);
            settings1.Append(rsids1);
            settings1.Append(mathProperties1);
            settings1.Append(themeFontLanguages1);
            settings1.Append(colorSchemeMapping1);
            settings1.Append(shapeDefaults1);
            settings1.Append(decimalSymbol1);
            settings1.Append(listSeparator1);

            documentSettingsPart1.Settings = settings1;
        }

        // Generates content of fontTablePart1.
        private void GenerateFontTablePart1Content(FontTablePart fontTablePart1)
        {
            Fonts fonts1 = new Fonts() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            fonts1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            fonts1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");

            Font font1 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number1 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "CC" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "C0007841", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font() { Name = "Calibri" };
            Panose1Number panose1Number2 = new Panose1Number() { Val = "020F0502020204030204" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "CC" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "4000ACFF", UnicodeSignature2 = "00000001", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font() { Name = "Arial Unicode MS" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "020B0604020202020204" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "80" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "F7FFAFFF", UnicodeSignature1 = "E9DFFFFF", UnicodeSignature2 = "0000003F", UnicodeSignature3 = "00000000", CodePageSignature0 = "003F01FF", CodePageSignature1 = "00000000" };

            font3.Append(panose1Number3);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            Font font4 = new Font() { Name = "Cambria" };
            Panose1Number panose1Number4 = new Panose1Number() { Val = "02040503050406030204" };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = "CC" };
            FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch4 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature4 = new FontSignature() { UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "400004FF", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font4.Append(panose1Number4);
            font4.Append(fontCharSet4);
            font4.Append(fontFamily4);
            font4.Append(pitch4);
            font4.Append(fontSignature4);

            Font font5 = new Font() { Name = "Segoe UI" };
            Panose1Number panose1Number5 = new Panose1Number() { Val = "020B0502040204020203" };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = "CC" };
            FontFamily fontFamily5 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch5 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature5 = new FontSignature() { UnicodeSignature0 = "E10022FF", UnicodeSignature1 = "C000E47F", UnicodeSignature2 = "00000029", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001DF", CodePageSignature1 = "00000000" };

            font5.Append(panose1Number5);
            font5.Append(fontCharSet5);
            font5.Append(fontFamily5);
            font5.Append(pitch5);
            font5.Append(fontSignature5);

            Font font6 = new Font() { Name = "Verdana" };
            Panose1Number panose1Number6 = new Panose1Number() { Val = "020B0604030504040204" };
            FontCharSet fontCharSet6 = new FontCharSet() { Val = "CC" };
            FontFamily fontFamily6 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch6 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature6 = new FontSignature() { UnicodeSignature0 = "A10006FF", UnicodeSignature1 = "4000205B", UnicodeSignature2 = "00000010", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font6.Append(panose1Number6);
            font6.Append(fontCharSet6);
            font6.Append(fontFamily6);
            font6.Append(pitch6);
            font6.Append(fontSignature6);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);

            fontTablePart1.Fonts = fonts1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Егоркин";
            document.PackageProperties.Revision = "2";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2014-10-10T12:57:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2014-10-10T12:57:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Atolkov";
        }


    }
}
