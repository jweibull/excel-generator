namespace TableExporter;

public static class StringSanitizer
{
    private static bool[]? _lookupTable;

    public static string RemoveSpecialCharacters(string str)
    {
        if (string.IsNullOrEmpty(str))
        {
            return string.Empty;
        }

        bool[] lookup = _lookupTable ??= BuildLookupTable();

        var sb = new StringBuilder(str.Length);
        foreach (char c in str)
        {
            if (c < lookup.Length && lookup[c])
            {
                sb.Append(c);
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Builds lookup table for valid XML 1.0 characters: Basic Latin block (U+0020–U+D7FF),
    /// Supplementary Private Use (U+E000–U+FFFD), and BMP beyond (U+10000–U+10FFFF), plus tab, LF, CR.
    /// </summary>
    private static bool[] BuildLookupTable()
    {
        var table = new bool[1114112];
        for (int c = 0x20; c <= 0xD7FF; c++) table[c] = true;
        for (int c = 0xE000; c <= 0xFFFD; c++) table[c] = true;
        for (int c = 0x10000; c <= 0x10FFFF; c++) table[c] = true;
        table[0x9] = true;  // tab
        table[0xA] = true;   // LF
        table[0xD] = true;   // CR
        return table;
    }
}
