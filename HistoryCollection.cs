using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class HistoryEntry
{
    public string? Year { get; set; }
    public string? Fields { get; set; }
    public string? Activities { get; set; }
}

public static class HistoryEntryTemplateMap
{
    public const string TemplateYear = "{{tplHistoryYear}}";
    public const string TemplateActivities = "{{tplHistoryActivities}}";
    public const string TemplateFields = "{{tplHistoryFields}}";

    // Maps property names of HistoryEntry to their corresponding template placeholders in the Word table
    public static readonly Dictionary<string, string> PropertyToTable = new Dictionary<string, string>
    {
        { nameof(HistoryEntry.Year), TemplateYear },
        { nameof(HistoryEntry.Activities), TemplateActivities },
        { nameof(HistoryEntry.Fields), TemplateFields }
    };

    // Maps template placeholders back to property names of HistoryEntry
    public static readonly Dictionary<string, string> TableToProperty = new Dictionary<string, string>
    {
        { TemplateYear, nameof(HistoryEntry.Year) },
        { TemplateActivities, nameof(HistoryEntry.Activities) },
        { TemplateFields, nameof(HistoryEntry.Fields) }
    };
}

public class HistoryCollection
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
    /// <param name="yamlFilePath">The path to the file containing the history entries.</param>
    /// <returns>A list of <see cref="HistoryEntry"/> objects parsed from the file.</returns>
    public static List<HistoryEntry> Deserialize(string yamlFilePath)
    {
        var entries = new List<HistoryEntry>();
        HistoryEntry? current = null;
        foreach (var line in File.ReadLines(yamlFilePath))
        {
            // Parse the YAML file 
            var trimmed = line.Trim();
            if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith("#"))
                continue;

            if (trimmed.StartsWith("- year:"))
            {
                if (current != null)
                    entries.Add(current);
                current = new HistoryEntry();
                var yearMatch = Regex.Match(trimmed, @"- year:\s*(\w+)");
                if (yearMatch.Success)
                    current.Year = yearMatch.Groups[1].Value;
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
                    if (current != null)
                        current.Fields = string.Join(", ", fieldsArray);
                }
            }
            else if (trimmed.StartsWith("activities:"))
            {
                // Remove ">-" if present
                var actLine = trimmed.Replace("activities:", "").Replace(">-", "").Trim();
                if (current != null)
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

    public static void ReplaceHistory(string docFilePath, List<HistoryEntry> historyItems)
    {
        //string tblCellMatchText = "{{tbl_history_year}}";
        // Get the table name for Year property
        HistoryEntryTemplateMap.PropertyToTable.TryGetValue(nameof(HistoryEntry.Year), out string? tblCellMatchText);

        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(docFilePath, true))
        {
            if (wordDoc.MainDocumentPart == null || wordDoc.MainDocumentPart.Document == null)
            {
                Console.WriteLine($"The document does not contain a main document part or document.");
                return;
            }

            var body = wordDoc.MainDocumentPart.Document.Body;
            if (body == null)
            {
                Console.WriteLine("The document body is null.");
                return;
            }

            // Find the original table by matching the content of cell (0,0)
            Table? templateTable = null;
            foreach (var table in body.Elements<Table>())
            {
                var firstRow = table.Elements<TableRow>().FirstOrDefault();
                var firstCell = firstRow?.Elements<TableCell>().FirstOrDefault();
                var cellText = firstCell?.InnerText;

                if (!string.IsNullOrEmpty(cellText) && !string.IsNullOrEmpty(tblCellMatchText) && cellText.Contains(tblCellMatchText))
                {
                    templateTable = table;
                    break;
                }
            }

            if (templateTable == null)
            {
                Console.WriteLine($"No table found with cell (0,0) text matching '{tblCellMatchText}'");
                return;
            }

            // Capture index of original table in its parent
            var parent = templateTable.Parent;
            if (parent == null)
            {
                Console.WriteLine("The template table's parent is null.");
                return;
            }
            int index = parent.ChildElements.ToList().IndexOf(templateTable);

            // Remove original table from the document
            templateTable.Remove();

            var tablesToInsert = new List<Table>();

            foreach (var item in historyItems)
            {
                // Clone the template table
                var newTable = (Table)templateTable.CloneNode(true);

                // Replace template placeholders in the cloned table
                foreach (var row in newTable.Elements<TableRow>())
                {
                    foreach (var cell in row.Elements<TableCell>())
                    {
                        // Gather all Text elements in this cell
                        var textElements = cell.Descendants<Text>().ToList();
                        if (textElements.Count < 3) continue;

                        // Look for the pattern: "*{{", someString, "}}"
                        for (int i = 0; i < textElements.Count - 2; i++)
                        {
                            string textStart = textElements[i].Text;
                            bool isStart = textStart.Length >= 2 && textStart.Substring(textStart.Length - 2) == "{{";
                            if (isStart && textElements[i + 2].Text == "}}")
                            {
                                string placeholder = "{{" + textElements[i + 1].Text + "}}";
                                if (HistoryEntryTemplateMap.TableToProperty.TryGetValue(placeholder, out string? propertyName))
                                {
                                    var property = typeof(HistoryEntry).GetProperty(propertyName);
                                    if (property != null)
                                    {
                                        // Replace the three tokens with the value, preserving formatting
                                        var value = property.GetValue(item)?.ToString() ?? "";
                                        // Assign "" to the last two characters of textElements[i].Text
                                        textElements[i].Text = textElements[i].Text.Substring(0, textElements[i].Text.Length - 2);
                                        textElements[i + 1].Text = value;
                                        textElements[i + 2].Text = "";
                                    }
                                }
                            }
                        }
                    }
                }

                tablesToInsert.Add(newTable);
            }

            // Insert all new tables at the original index
            int insertIndex = index;
            foreach (var tbl in tablesToInsert)
            {
                parent.InsertAt(tbl, insertIndex++);
                // Optionally add a blank paragraph between tables
                var emptyParagraph = new Paragraph(new Run(new Text("")));
                parent.InsertAt(emptyParagraph, insertIndex++);
            }

/*
            // Clone templateTable
            Table newTableContent = templateTable;

            // Insert the new table at the same index
            parent.InsertAt(newTableContent, index);

            index = AppendTable(index, parent, newTableContent);
            index = AppendTable(index, parent, newTableContent);
*/

            // Save both the document and its main part
            wordDoc.MainDocumentPart.Document.Save();
            wordDoc.Save();
        }
    }
}

