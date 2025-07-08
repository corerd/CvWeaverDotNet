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
        // Get the table placeholder from Year property
        HistoryEntryTemplateMap.PropertyToTable.TryGetValue(nameof(HistoryEntry.Year), out string? tblCellMatchText);
        if (string.IsNullOrEmpty(tblCellMatchText))
        {
            Console.WriteLine("No table placeholder found.");
            return;
        }

        var template = DataCollection.ExtractTableAtPlaceholder(docxBody, tblCellMatchText);
        if (template.FoundTable == null || template.Parent == null || template.Index < 0)
            return;

        var tablesToInsert = new List<Table>();

        foreach (var item in historyItems)
        {
            // Clone the template table
            var newTable = (Table)template.FoundTable.CloneNode(true);

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
        int insertIndex = template.Index;
        foreach (var tbl in tablesToInsert)
        {
            template.Parent.InsertAt(tbl, insertIndex++);
            // Optionally add a blank paragraph between tables
            var emptyParagraph = new Paragraph(new Run(new Text("")));
            template.Parent.InsertAt(emptyParagraph, insertIndex++);
        }
    }

    public static void MergeHistoryData(Body docxBody, string dataSetFilePath)
    {
        List<HistoryEntry> historyItems = DataCollection.DeserializeYAML<HistoryEntry>(dataSetFilePath);
        ReplaceHistoryTemplate(docxBody, historyItems);
    }
}

