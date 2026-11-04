using System.IO.Compression;
using System.Text;

namespace XlWright.McpServer.Tests;

// Writes the smallest SpreadsheetML package Excel opens: one empty sheet, and optionally a second, protected one
// named Locked. The package gives no CodeName, so Excel assigns one. No Excel is needed to make it.
internal static class MinimalWorkbook
{
    public static void Write(string path, bool lockedSheet = false, string firstSheet = "Sheet1")
    {
        var macroEnabled = Path.GetExtension(path).Equals(".xlsm", StringComparison.OrdinalIgnoreCase);
        var mainType = macroEnabled
            ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
            : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        const string Relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string Office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string WorksheetType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(zip, "[Content_Types].xml",
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            $"<Override PartName=\"/xl/workbook.xml\" ContentType=\"{mainType}\"/>" +
            $"<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"{WorksheetType}\"/>" +
            (lockedSheet ? $"<Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"{WorksheetType}\"/>" : "") +
            "</Types>");
        Add(zip, "_rels/.rels",
            $"<Relationships xmlns=\"{Relationships}\"><Relationship Id=\"rId1\" Type=\"{Office}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
        Add(zip, "xl/workbook.xml",
            $"<workbook xmlns=\"{Main}\" xmlns:r=\"{Office}\"><sheets><sheet name=\"{firstSheet}\" sheetId=\"1\" r:id=\"rId1\"/>" +
            (lockedSheet ? "<sheet name=\"Locked\" sheetId=\"2\" r:id=\"rId2\"/>" : "") +
            "</sheets></workbook>");
        Add(zip, "xl/_rels/workbook.xml.rels",
            $"<Relationships xmlns=\"{Relationships}\"><Relationship Id=\"rId1\" Type=\"{Office}/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
            (lockedSheet ? $"<Relationship Id=\"rId2\" Type=\"{Office}/worksheet\" Target=\"worksheets/sheet2.xml\"/>" : "") +
            "</Relationships>");
        Add(zip, "xl/worksheets/sheet1.xml", $"<worksheet xmlns=\"{Main}\"><sheetData/></worksheet>");
        if (lockedSheet)
        {
            Add(zip, "xl/worksheets/sheet2.xml", $"<worksheet xmlns=\"{Main}\"><sheetData/><sheetProtection sheet=\"1\"/></worksheet>");
        }
    }

    private static void Add(ZipArchive zip, string name, string xml)
    {
        using var stream = zip.CreateEntry(name).Open();
        var bytes = Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + xml);
        stream.Write(bytes);
    }
}
