using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X15ac = DocumentFormat.OpenXml.Office2013.ExcelAc;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;

namespace ExcelGenerator.Generators;

public class MinimalBoilerPlate
{
    // Creates a SpreadsheetDocument.
    public void CreatePackage(string filePath)
    {
        using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
        {
            CreateParts(package);
        }
    }

    // Adds child parts and generates content of the specified part.
    private void CreateParts(SpreadsheetDocument document)
    {
        WorkbookPart workbookPart1 = document.AddWorkbookPart();
        GenerateWorkbookPart1Content(workbookPart1);

        ThemePart themePart1 = workbookPart1.AddNewPart<ThemePart>("rId1");
        GenerateThemePart1Content(themePart1);

        WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId2");
        GenerateWorkbookStylesPart1Content(workbookStylesPart1);

        SharedStringTablePart sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId3");
        GenerateSharedStringTablePart1Content(sharedStringTablePart1);

        WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId4");
        GenerateWorksheetPart1Content(worksheetPart1);

        TableDefinitionPart tableDefinitionPart1 = worksheetPart1.AddNewPart<TableDefinitionPart>("rId3");
        GenerateTableDefinitionPart1Content(tableDefinitionPart1);

        //Metadata
        SetPackageProperties(document);
    }

    // Generates content of workbookPart1.
    private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
    {
        Workbook workbook1 = new Workbook();
        workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        workbook1.AddNamespaceDeclaration("mx", "http://schemas.microsoft.com/office/mac/excel/2008/main");
        workbook1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        workbook1.AddNamespaceDeclaration("mv", "urn:schemas-microsoft-com:mac:vml");
        workbook1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
        workbook1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
        workbook1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
        workbook1.AddNamespaceDeclaration("xm", "http://schemas.microsoft.com/office/excel/2006/main");
        WorkbookProperties workbookProperties1 = new WorkbookProperties();

        Sheets sheets1 = new Sheets();
        Sheet sheet1 = new Sheet() { Name = "Página1", SheetId = (UInt32Value)1U, State = SheetStateValues.Visible, Id = "rId4" };

        sheets1.Append(sheet1);
        DefinedNames definedNames1 = new DefinedNames();
        CalculationProperties calculationProperties1 = new CalculationProperties();

        workbook1.Append(workbookProperties1);
        workbook1.Append(sheets1);
        workbook1.Append(definedNames1);
        workbook1.Append(calculationProperties1);

        workbookPart1.Workbook = workbook1;
    }

    // Generates content of themePart1.
    private void GenerateThemePart1Content(ThemePart themePart1)
    {
        A.Theme theme1 = new A.Theme() { Name = "Sheets" };
        theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        theme1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        A.ThemeElements themeElements1 = new A.ThemeElements();

        A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Sheets" };

        A.Dark1Color dark1Color1 = new A.Dark1Color();
        A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "000000" };

        dark1Color1.Append(rgbColorModelHex1);

        A.Light1Color light1Color1 = new A.Light1Color();
        A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "FFFFFF" };

        light1Color1.Append(rgbColorModelHex2);

        A.Dark2Color dark2Color1 = new A.Dark2Color();
        A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "000000" };

        dark2Color1.Append(rgbColorModelHex3);

        A.Light2Color light2Color1 = new A.Light2Color();
        A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "FFFFFF" };

        light2Color1.Append(rgbColorModelHex4);

        A.Accent1Color accent1Color1 = new A.Accent1Color();
        A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "4285F4" };

        accent1Color1.Append(rgbColorModelHex5);

        A.Accent2Color accent2Color1 = new A.Accent2Color();
        A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "EA4335" };

        accent2Color1.Append(rgbColorModelHex6);

        A.Accent3Color accent3Color1 = new A.Accent3Color();
        A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "FBBC04" };

        accent3Color1.Append(rgbColorModelHex7);

        A.Accent4Color accent4Color1 = new A.Accent4Color();
        A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "34A853" };

        accent4Color1.Append(rgbColorModelHex8);

        A.Accent5Color accent5Color1 = new A.Accent5Color();
        A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "FF6D01" };

        accent5Color1.Append(rgbColorModelHex9);

        A.Accent6Color accent6Color1 = new A.Accent6Color();
        A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "46BDC6" };

        accent6Color1.Append(rgbColorModelHex10);

        A.Hyperlink hyperlink1 = new A.Hyperlink();
        A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "1155CC" };

        hyperlink1.Append(rgbColorModelHex11);

        A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
        A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "1155CC" };

        followedHyperlinkColor1.Append(rgbColorModelHex12);

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
        colorScheme1.Append(hyperlink1);
        colorScheme1.Append(followedHyperlinkColor1);

        A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Sheets" };

        A.MajorFont majorFont1 = new A.MajorFont();
        A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Arial" };
        A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "Arial" };
        A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "Arial" };

        majorFont1.Append(latinFont1);
        majorFont1.Append(eastAsianFont1);
        majorFont1.Append(complexScriptFont1);

        A.MinorFont minorFont1 = new A.MinorFont();
        A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Arial" };
        A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "Arial" };
        A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "Arial" };

        minorFont1.Append(latinFont2);
        minorFont1.Append(eastAsianFont2);
        minorFont1.Append(complexScriptFont2);

        fontScheme1.Append(majorFont1);
        fontScheme1.Append(minorFont1);

        A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

        A.FillStyleList fillStyleList1 = new A.FillStyleList();

        A.SolidFill solidFill1 = new A.SolidFill();
        A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

        solidFill1.Append(schemeColor1);

        A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

        A.GradientStopList gradientStopList1 = new A.GradientStopList();

        A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

        A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 110000 };
        A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
        A.Tint tint1 = new A.Tint() { Val = 67000 };

        schemeColor2.Append(luminanceModulation1);
        schemeColor2.Append(saturationModulation1);
        schemeColor2.Append(tint1);

        gradientStop1.Append(schemeColor2);

        A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

        A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
        A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
        A.Tint tint2 = new A.Tint() { Val = 73000 };

        schemeColor3.Append(luminanceModulation2);
        schemeColor3.Append(saturationModulation2);
        schemeColor3.Append(tint2);

        gradientStop2.Append(schemeColor3);

        A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

        A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
        A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
        A.Tint tint3 = new A.Tint() { Val = 81000 };

        schemeColor4.Append(luminanceModulation3);
        schemeColor4.Append(saturationModulation3);
        schemeColor4.Append(tint3);

        gradientStop3.Append(schemeColor4);

        gradientStopList1.Append(gradientStop1);
        gradientStopList1.Append(gradientStop2);
        gradientStopList1.Append(gradientStop3);
        A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

        gradientFill1.Append(gradientStopList1);
        gradientFill1.Append(linearGradientFill1);

        A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

        A.GradientStopList gradientStopList2 = new A.GradientStopList();

        A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

        A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
        A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
        A.Tint tint4 = new A.Tint() { Val = 94000 };

        schemeColor5.Append(saturationModulation4);
        schemeColor5.Append(luminanceModulation4);
        schemeColor5.Append(tint4);

        gradientStop4.Append(schemeColor5);

        A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

        A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
        A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
        A.Shade shade1 = new A.Shade() { Val = 100000 };

        schemeColor6.Append(saturationModulation5);
        schemeColor6.Append(luminanceModulation5);
        schemeColor6.Append(shade1);

        gradientStop5.Append(schemeColor6);

        A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

        A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
        A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
        A.Shade shade2 = new A.Shade() { Val = 78000 };

        schemeColor7.Append(luminanceModulation6);
        schemeColor7.Append(saturationModulation6);
        schemeColor7.Append(shade2);

        gradientStop6.Append(schemeColor7);

        gradientStopList2.Append(gradientStop4);
        gradientStopList2.Append(gradientStop5);
        gradientStopList2.Append(gradientStop6);
        A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

        gradientFill2.Append(gradientStopList2);
        gradientFill2.Append(linearGradientFill2);

        fillStyleList1.Append(solidFill1);
        fillStyleList1.Append(gradientFill1);
        fillStyleList1.Append(gradientFill2);

        A.LineStyleList lineStyleList1 = new A.LineStyleList();

        A.Outline outline1 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

        A.SolidFill solidFill2 = new A.SolidFill();
        A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

        solidFill2.Append(schemeColor8);
        A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
        A.Miter miter1 = new A.Miter() { Limit = 800000 };

        outline1.Append(solidFill2);
        outline1.Append(presetDash1);
        outline1.Append(miter1);

        A.Outline outline2 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

        A.SolidFill solidFill3 = new A.SolidFill();
        A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

        solidFill3.Append(schemeColor9);
        A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
        A.Miter miter2 = new A.Miter() { Limit = 800000 };

        outline2.Append(solidFill3);
        outline2.Append(presetDash2);
        outline2.Append(miter2);

        A.Outline outline3 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

        A.SolidFill solidFill4 = new A.SolidFill();
        A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

        solidFill4.Append(schemeColor10);
        A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
        A.Miter miter3 = new A.Miter() { Limit = 800000 };

        outline3.Append(solidFill4);
        outline3.Append(presetDash3);
        outline3.Append(miter3);

        lineStyleList1.Append(outline1);
        lineStyleList1.Append(outline2);
        lineStyleList1.Append(outline3);

        A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

        A.EffectStyle effectStyle1 = new A.EffectStyle();
        A.EffectList effectList1 = new A.EffectList();

        effectStyle1.Append(effectList1);

        A.EffectStyle effectStyle2 = new A.EffectStyle();
        A.EffectList effectList2 = new A.EffectList();

        effectStyle2.Append(effectList2);

        A.EffectStyle effectStyle3 = new A.EffectStyle();

        A.EffectList effectList3 = new A.EffectList();

        A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

        A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
        A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

        rgbColorModelHex13.Append(alpha1);

        outerShadow1.Append(rgbColorModelHex13);

        effectList3.Append(outerShadow1);

        effectStyle3.Append(effectList3);

        effectStyleList1.Append(effectStyle1);
        effectStyleList1.Append(effectStyle2);
        effectStyleList1.Append(effectStyle3);

        A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

        A.SolidFill solidFill5 = new A.SolidFill();
        A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

        solidFill5.Append(schemeColor11);

        A.SolidFill solidFill6 = new A.SolidFill();

        A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        A.Tint tint5 = new A.Tint() { Val = 95000 };
        A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

        schemeColor12.Append(tint5);
        schemeColor12.Append(saturationModulation7);

        solidFill6.Append(schemeColor12);

        A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

        A.GradientStopList gradientStopList3 = new A.GradientStopList();

        A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

        A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        A.Tint tint6 = new A.Tint() { Val = 93000 };
        A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
        A.Shade shade3 = new A.Shade() { Val = 98000 };
        A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

        schemeColor13.Append(tint6);
        schemeColor13.Append(saturationModulation8);
        schemeColor13.Append(shade3);
        schemeColor13.Append(luminanceModulation7);

        gradientStop7.Append(schemeColor13);

        A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

        A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        A.Tint tint7 = new A.Tint() { Val = 98000 };
        A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
        A.Shade shade4 = new A.Shade() { Val = 90000 };
        A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

        schemeColor14.Append(tint7);
        schemeColor14.Append(saturationModulation9);
        schemeColor14.Append(shade4);
        schemeColor14.Append(luminanceModulation8);

        gradientStop8.Append(schemeColor14);

        A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

        A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        A.Shade shade5 = new A.Shade() { Val = 63000 };
        A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

        schemeColor15.Append(shade5);
        schemeColor15.Append(saturationModulation10);

        gradientStop9.Append(schemeColor15);

        gradientStopList3.Append(gradientStop7);
        gradientStopList3.Append(gradientStop8);
        gradientStopList3.Append(gradientStop9);
        A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

        gradientFill3.Append(gradientStopList3);
        gradientFill3.Append(linearGradientFill3);

        backgroundFillStyleList1.Append(solidFill5);
        backgroundFillStyleList1.Append(solidFill6);
        backgroundFillStyleList1.Append(gradientFill3);

        formatScheme1.Append(fillStyleList1);
        formatScheme1.Append(lineStyleList1);
        formatScheme1.Append(effectStyleList1);
        formatScheme1.Append(backgroundFillStyleList1);

        themeElements1.Append(colorScheme1);
        themeElements1.Append(fontScheme1);
        themeElements1.Append(formatScheme1);

        theme1.Append(themeElements1);

        themePart1.Theme = theme1;
    }

    // Generates content of workbookStylesPart1.
    private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
    {
        Stylesheet stylesheet1 = new Stylesheet();
        stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
        stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

        Fonts fonts1 = new Fonts() { Count = (UInt32Value)2U };

        Font font1 = new Font();
        FontSize fontSize1 = new FontSize() { Val = 10.0D };
        Color color1 = new Color() { Rgb = "FF000000" };
        FontName fontName1 = new FontName() { Val = "Arial" };
        FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

        font1.Append(fontSize1);
        font1.Append(color1);
        font1.Append(fontName1);
        font1.Append(fontScheme2);

        Font font2 = new Font();
        Color color2 = new Color() { Theme = (UInt32Value)1U };
        FontName fontName2 = new FontName() { Val = "Arial" };
        FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Minor };

        font2.Append(color2);
        font2.Append(fontName2);
        font2.Append(fontScheme3);

        fonts1.Append(font1);
        fonts1.Append(font2);

        Fills fills1 = new Fills() { Count = (UInt32Value)2U };

        Fill fill1 = new Fill();
        PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

        fill1.Append(patternFill1);

        Fill fill2 = new Fill();
        PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.LightGray };

        fill2.Append(patternFill2);

        fills1.Append(fill1);
        fills1.Append(fill2);

        Borders borders1 = new Borders() { Count = (UInt32Value)1U };
        Border border1 = new Border();

        borders1.Append(border1);

        CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
        CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };

        cellStyleFormats1.Append(cellFormat1);

        CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)2U };

        CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
        Alignment alignment1 = new Alignment() { Vertical = VerticalAlignmentValues.Bottom, WrapText = false, ShrinkToFit = false, ReadingOrder = (UInt32Value)0U };

        cellFormat2.Append(alignment1);

        CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
        Alignment alignment2 = new Alignment() { ReadingOrder = (UInt32Value)0U };

        cellFormat3.Append(alignment2);

        cellFormats1.Append(cellFormat2);
        cellFormats1.Append(cellFormat3);

        CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
        CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

        cellStyles1.Append(cellStyle1);

        DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)4U };

        DifferentialFormat differentialFormat1 = new DifferentialFormat();
        Font font3 = new Font();

        Fill fill3 = new Fill();
        PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.None };

        fill3.Append(patternFill3);
        Border border2 = new Border();

        differentialFormat1.Append(font3);
        differentialFormat1.Append(fill3);
        differentialFormat1.Append(border2);

        DifferentialFormat differentialFormat2 = new DifferentialFormat();
        Font font4 = new Font();

        Fill fill4 = new Fill();

        PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "FFBDBDBD" };
        BackgroundColor backgroundColor1 = new BackgroundColor() { Rgb = "FFBDBDBD" };

        patternFill4.Append(foregroundColor1);
        patternFill4.Append(backgroundColor1);

        fill4.Append(patternFill4);
        Border border3 = new Border();

        differentialFormat2.Append(font4);
        differentialFormat2.Append(fill4);
        differentialFormat2.Append(border3);

        DifferentialFormat differentialFormat3 = new DifferentialFormat();
        Font font5 = new Font();

        Fill fill5 = new Fill();

        PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = "FFFFFFFF" };
        BackgroundColor backgroundColor2 = new BackgroundColor() { Rgb = "FFFFFFFF" };

        patternFill5.Append(foregroundColor2);
        patternFill5.Append(backgroundColor2);

        fill5.Append(patternFill5);
        Border border4 = new Border();

        differentialFormat3.Append(font5);
        differentialFormat3.Append(fill5);
        differentialFormat3.Append(border4);

        DifferentialFormat differentialFormat4 = new DifferentialFormat();
        Font font6 = new Font();

        Fill fill6 = new Fill();

        PatternFill patternFill6 = new PatternFill() { PatternType = PatternValues.Solid };
        ForegroundColor foregroundColor3 = new ForegroundColor() { Rgb = "FFF3F3F3" };
        BackgroundColor backgroundColor3 = new BackgroundColor() { Rgb = "FFF3F3F3" };

        patternFill6.Append(foregroundColor3);
        patternFill6.Append(backgroundColor3);

        fill6.Append(patternFill6);
        Border border5 = new Border();

        differentialFormat4.Append(font6);
        differentialFormat4.Append(fill6);
        differentialFormat4.Append(border5);

        differentialFormats1.Append(differentialFormat1);
        differentialFormats1.Append(differentialFormat2);
        differentialFormats1.Append(differentialFormat3);
        differentialFormats1.Append(differentialFormat4);

        TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)1U };

        TableStyle tableStyle1 = new TableStyle() { Name = "Página1-style", Pivot = false, Count = (UInt32Value)3U };
        TableStyleElement tableStyleElement1 = new TableStyleElement() { Type = TableStyleValues.HeaderRow, FormatId = (UInt32Value)1U };
        TableStyleElement tableStyleElement2 = new TableStyleElement() { Type = TableStyleValues.FirstRowStripe, FormatId = (UInt32Value)2U };
        TableStyleElement tableStyleElement3 = new TableStyleElement() { Type = TableStyleValues.SecondRowStripe, FormatId = (UInt32Value)3U };

        tableStyle1.Append(tableStyleElement1);
        tableStyle1.Append(tableStyleElement2);
        tableStyle1.Append(tableStyleElement3);

        tableStyles1.Append(tableStyle1);

        stylesheet1.Append(fonts1);
        stylesheet1.Append(fills1);
        stylesheet1.Append(borders1);
        stylesheet1.Append(cellStyleFormats1);
        stylesheet1.Append(cellFormats1);
        stylesheet1.Append(cellStyles1);
        stylesheet1.Append(differentialFormats1);
        stylesheet1.Append(tableStyles1);

        workbookStylesPart1.Stylesheet = stylesheet1;
    }

    // Generates content of sharedStringTablePart1.
    private void GenerateSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1)
    {
        SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)6U, UniqueCount = (UInt32Value)6U };

        SharedStringItem sharedStringItem1 = new SharedStringItem();
        Text text1 = new Text();
        text1.Text = "Header 1";

        sharedStringItem1.Append(text1);

        SharedStringItem sharedStringItem2 = new SharedStringItem();
        Text text2 = new Text();
        text2.Text = "Header 2";

        sharedStringItem2.Append(text2);

        SharedStringItem sharedStringItem3 = new SharedStringItem();
        Text text3 = new Text();
        text3.Text = "Cell A2";

        sharedStringItem3.Append(text3);

        SharedStringItem sharedStringItem4 = new SharedStringItem();
        Text text4 = new Text();
        text4.Text = "Cell B2";

        sharedStringItem4.Append(text4);

        SharedStringItem sharedStringItem5 = new SharedStringItem();
        Text text5 = new Text();
        text5.Text = "Cell A3";

        sharedStringItem5.Append(text5);

        SharedStringItem sharedStringItem6 = new SharedStringItem();
        Text text6 = new Text();
        text6.Text = "CellB3";

        sharedStringItem6.Append(text6);

        sharedStringTable1.Append(sharedStringItem1);
        sharedStringTable1.Append(sharedStringItem2);
        sharedStringTable1.Append(sharedStringItem3);
        sharedStringTable1.Append(sharedStringItem4);
        sharedStringTable1.Append(sharedStringItem5);
        sharedStringTable1.Append(sharedStringItem6);

        sharedStringTablePart1.SharedStringTable = sharedStringTable1;
    }

    // Generates content of worksheetPart1.
    private void GenerateWorksheetPart1Content(WorksheetPart worksheetPart1)
    {
        Worksheet worksheet1 = new Worksheet();
        worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        worksheet1.AddNamespaceDeclaration("mx", "http://schemas.microsoft.com/office/mac/excel/2008/main");
        worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        worksheet1.AddNamespaceDeclaration("mv", "urn:schemas-microsoft-com:mac:vml");
        worksheet1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
        worksheet1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
        worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
        worksheet1.AddNamespaceDeclaration("xm", "http://schemas.microsoft.com/office/excel/2006/main");

        SheetProperties sheetProperties1 = new SheetProperties();
        OutlineProperties outlineProperties1 = new OutlineProperties() { SummaryBelow = false, SummaryRight = false };

        sheetProperties1.Append(outlineProperties1);

        SheetViews sheetViews1 = new SheetViews();
        SheetView sheetView1 = new SheetView() { WorkbookViewId = (UInt32Value)0U };

        sheetViews1.Append(sheetView1);
        SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultColumnWidth = 12.63D, DefaultRowHeight = 15.75D, CustomHeight = true };

        SheetData sheetData1 = new SheetData();

        Row row1 = new Row() { RowIndex = (UInt32Value)1U };

        Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
        CellValue cellValue1 = new CellValue();
        cellValue1.Text = "0";

        cell1.Append(cellValue1);

        Cell cell2 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
        CellValue cellValue2 = new CellValue();
        cellValue2.Text = "1";

        cell2.Append(cellValue2);

        row1.Append(cell1);
        row1.Append(cell2);

        Row row2 = new Row() { RowIndex = (UInt32Value)2U };

        Cell cell3 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
        CellValue cellValue3 = new CellValue();
        cellValue3.Text = "2";

        cell3.Append(cellValue3);

        Cell cell4 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
        CellValue cellValue4 = new CellValue();
        cellValue4.Text = "3";

        cell4.Append(cellValue4);

        row2.Append(cell3);
        row2.Append(cell4);

        Row row3 = new Row() { RowIndex = (UInt32Value)3U };

        Cell cell5 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
        CellValue cellValue5 = new CellValue();
        cellValue5.Text = "4";

        cell5.Append(cellValue5);

        Cell cell6 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
        CellValue cellValue6 = new CellValue();
        cellValue6.Text = "5";

        cell6.Append(cellValue6);

        row3.Append(cell5);
        row3.Append(cell6);

        sheetData1.Append(row1);
        sheetData1.Append(row2);
        sheetData1.Append(row3);
        
        TableParts tableParts1 = new TableParts() { Count = (UInt32Value)1U };
        TablePart tablePart1 = new TablePart() { Id = "rId3" };

        tableParts1.Append(tablePart1);

        worksheet1.Append(sheetProperties1);
        worksheet1.Append(sheetViews1);
        worksheet1.Append(sheetFormatProperties1);
        worksheet1.Append(sheetData1);
        worksheet1.Append(tableParts1);

        worksheetPart1.Worksheet = worksheet1;
    }

    // Generates content of tableDefinitionPart1.
    private void GenerateTableDefinitionPart1Content(TableDefinitionPart tableDefinitionPart1)
    {
        Table table1 = new Table() { Id = (UInt32Value)1U, DisplayName = "Table_1", Reference = "A1:B3" };

        TableColumns tableColumns1 = new TableColumns() { Count = (UInt32Value)2U };
        TableColumn tableColumn1 = new TableColumn() { Id = (UInt32Value)1U, Name = "Header 1" };
        TableColumn tableColumn2 = new TableColumn() { Id = (UInt32Value)2U, Name = "Header 2" };

        tableColumns1.Append(tableColumn1);
        tableColumns1.Append(tableColumn2);
        TableStyleInfo tableStyleInfo1 = new TableStyleInfo() { Name = "Página1-style", ShowFirstColumn = true, ShowLastColumn = true, ShowRowStripes = true, ShowColumnStripes = false };

        table1.Append(tableColumns1);
        table1.Append(tableStyleInfo1);

        tableDefinitionPart1.Table = table1;
    }

    private void SetPackageProperties(OpenXmlPackage document)
    {
    }


}

