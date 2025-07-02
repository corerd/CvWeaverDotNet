using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

public class HistoryEntry
{
    public int Year { get; set; }
    public string Fields { get; set; }
    public string Activities { get; set; }
}

public class DataCollection
{
    /// <summary>
    /// Reads history entries from a specified file path and parses them into a list of <see cref="HistoryEntry"/> objects.
    /// The file is expected to have entries in a specific format, where each entry starts with "- year:",
    /// followed by optional "fields:" and "activities:" lines.
    /// Lines starting with "#" (comment) or empty lines are ignored.
    /// Multiline "activities" are supported by means of the YAML >- block scalar indicator.
    /// Example:
    /// # This is a comment line
    ///      <this is an empty line>
    /// - year: XYWZ
    ///   fields: [item_1, ..., item_n]
    ///   activities: >-
    ///     This is a Multiline activity description following the YAML >- block scalar indicator.
    /// </summary>
    /// <param name="path">The path to the file containing the history entries.</param>
    /// <returns>A list of <see cref="HistoryEntry"/> objects parsed from the file.</returns>
    public static List<HistoryEntry> ReadHistoryEntries(string path)
    {
        var entries = new List<HistoryEntry>();
        HistoryEntry current = null;
        foreach (var line in File.ReadLines(path))
        {
            var trimmed = line.Trim();
            if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith("#"))
                continue;

            if (trimmed.StartsWith("- year:"))
            {
                if (current != null)
                    entries.Add(current);
                current = new HistoryEntry();
                var yearMatch = Regex.Match(trimmed, @"- year:\s*(\d+)");
                if (yearMatch.Success)
                    current.Year = int.Parse(yearMatch.Groups[1].Value);
            }
            else if (trimmed.StartsWith("fields:"))
            {
                var fieldsMatch = Regex.Match(trimmed, @"fields:\s*\[(.*?)\]");
                if (fieldsMatch.Success)
                {
                    // Remove spaces after commas and trim each item
                    var fieldsRaw = fieldsMatch.Groups[1].Value;
                    var fieldsArray = fieldsRaw.Split(',');
                    for (int i = 0; i < fieldsArray.Length; i++)
                        fieldsArray[i] = fieldsArray[i].Trim();
                    current.Fields = string.Join(", ", fieldsArray);
                }
            }
            else if (trimmed.StartsWith("activities:"))
            {
                // Remove ">-" if present
                var actLine = trimmed.Replace("activities:", "").Replace(">-", "").Trim();
                current.Activities = actLine;
            }
            else if (current != null && !trimmed.StartsWith("- year:") && !trimmed.StartsWith("fields:") && !trimmed.StartsWith("activities:"))
            {
                // Multiline activities
                if (!string.IsNullOrEmpty(current.Activities))
                    current.Activities += " ";
                current.Activities += trimmed;
            }
        }
        if (current != null)
            entries.Add(current);
        return entries;
    }
}
