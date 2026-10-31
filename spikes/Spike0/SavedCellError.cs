using System.Globalization;
using System.Xml.Linq;

namespace Spike0;

// Only the rich error forms observed in S4. Unknown/malformed metadata must fail verification.
internal static class SavedCellError
{
    internal static string Read(XElement cell, Func<string, XDocument> read)
    {
        XNamespace sheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace rich = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
        string value = cell.Element(sheet + "v")?.Value ?? throw new InvalidDataException("Missing saved error value.");
        if (cell.Attribute("vm") is null) return value;
        CellProbe.Require(value == "#VALUE!", "rich error compatibility value");
        var metadata = read("xl/metadata.xml").Root ?? throw new InvalidDataException("Missing metadata root.");
        var reference = At(metadata.Element(sheet + "valueMetadata")?.Elements(sheet + "bk"),
            Index(cell.Attribute("vm"), 1)).Elements(sheet + "rc").Single();
        var type = At(metadata.Element(sheet + "metadataTypes")?.Elements(sheet + "metadataType"),
            Index(reference.Attribute("t"), 1));
        CellProbe.Require((string?)type.Attribute("name") == "XLRICHVALUE", "rich metadata type");
        var future = metadata.Elements(sheet + "futureMetadata").Single(item => (string?)item.Attribute("name") == "XLRICHVALUE");
        var richReference = At(future.Elements(sheet + "bk"), Index(reference.Attribute("v"), 0))
            .Descendants(rich + "rvb").Single();
        var data = read("xl/richData/rdrichvalue.xml").Root ?? throw new InvalidDataException("Missing rich values.");
        var record = At(data.Elements(rich + "rv"), Index(richReference.Attribute("i"), 0));
        var structures = read("xl/richData/rdrichvaluestructure.xml").Root ?? throw new InvalidDataException("Missing rich structures.");
        var structure = At(structures.Elements(rich + "s"), Index(record.Attribute("s"), 0));
        CellProbe.Require((string?)structure.Attribute("t") == "_error", "rich structure type");
        var keys = structure.Elements(rich + "k").ToArray();
        var values = record.Elements(rich + "v").ToArray();
        CellProbe.Require(keys.Length == values.Length && keys.Select(key => (string?)key.Attribute("n")).Distinct().Count() == keys.Length,
            "rich error keys/values");
        int errorIndex = Array.FindIndex(keys, key => (string?)key.Attribute("n") == "errorType");
        CellProbe.Require(errorIndex >= 0 && (string?)keys[errorIndex].Attribute("t") == "i", "rich errorType key");
        return values[errorIndex].Value switch
        {
            "8" => "#SPILL!", "13" => "#CALC!",
            _ => throw new InvalidDataException("Unsupported rich error type; preserve evidence.")
        };
    }

    private static int Index(XAttribute? attribute, int basis)
    {
        if (!int.TryParse(attribute?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out int value) || value < basis)
            throw new InvalidDataException("Invalid rich metadata index.");
        return value - basis;
    }

    private static XElement At(IEnumerable<XElement>? elements, int index) =>
        elements?.ElementAtOrDefault(index) ?? throw new InvalidDataException("Rich metadata reference is missing or out of range.");
}
