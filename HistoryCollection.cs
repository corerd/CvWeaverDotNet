using DocumentFormat.OpenXml.Wordprocessing;

public class HistoryEntry
{
    public string? Year { get; set; }
    public List<string>? Fields { get; set; }
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
    public static void ReplaceHistoryTemplate(Body docxBody, List<HistoryEntry> historyItems)
    {
        //string tblCellMatchText = "{{tbl_history_year}}";
        // Get the table name for Year property
        HistoryEntryTemplateMap.PropertyToTable.TryGetValue(nameof(HistoryEntry.Year), out string? tblCellMatchText);

        // Find the original table by matching the content of cell (0,0)
        Table? templateTable = null;
        foreach (var table in docxBody.Elements<Table>())
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
                                    string propertyValueString = "";
                                    if (property.PropertyType == typeof(List<string>))
                                    {
                                        List<string>? fieldsList = property.GetValue(item) as List<string>;
                                        propertyValueString = string.Join(", ", fieldsList ?? new List<string>());
                                    }
                                    else
                                    {
                                        propertyValueString = property.GetValue(item)?.ToString() ?? "";
                                    }
                                    // Replace the three tokens with the value, preserving formatting
                                    // textElements[i].Text is "*{{", then remove the last two characters
                                    textElements[i].Text = textElements[i].Text.Substring(0, textElements[i].Text.Length - 2);
                                    textElements[i + 1].Text = propertyValueString;
                                    // Assign "" to the last two characters of textElements[i + 2].Text
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
    }

    public static void MergeHistoryData(Body docxBody, string dataSetFilePath)
    {
        List<HistoryEntry> historyItems = DataCollection.DeserializeYAML<HistoryEntry>(dataSetFilePath);
        ReplaceHistoryTemplate(docxBody, historyItems);
    }
}

