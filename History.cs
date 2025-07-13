using DocumentFormat.OpenXml.Wordprocessing;

public class History
{
    private class YamlEntry
    {
        public string? Year { get; set; }
        public List<string>? Fields { get; set; }
        public string? Activities { get; set; }
    }

    private static class PlaceholderYamlPropertyMap
    {
        public const string TemplateYear = "{{tplHistoryYear}}";
        public const string TemplateActivities = "{{tplHistoryActivities}}";
        public const string TemplateFields = "{{tplHistoryFields}}";

        // Maps property names of HistoryEntry to their corresponding template placeholders in the Word table
        public static readonly Dictionary<string, string> PropertyToTable = new Dictionary<string, string>
        {
            { nameof(YamlEntry.Year), TemplateYear },
            { nameof(YamlEntry.Activities), TemplateActivities },
            { nameof(YamlEntry.Fields), TemplateFields }
        };

        // Maps template placeholders back to property names of HistoryEntry
        public static readonly Dictionary<string, string> TableToProperty = new Dictionary<string, string>
        {
            { TemplateYear, nameof(YamlEntry.Year) },
            { TemplateActivities, nameof(YamlEntry.Activities) },
            { TemplateFields, nameof(YamlEntry.Fields) }
        };
    }

    private static void ReplaceWordTemplate(Body docxBody, List<YamlEntry> dataList)
    {
        // Get the table placeholder from Year property
        PlaceholderYamlPropertyMap.PropertyToTable.TryGetValue(nameof(YamlEntry.Year), out string? tblCellMatchText);
        if (string.IsNullOrEmpty(tblCellMatchText))
        {
            Console.WriteLine("No table placeholder found.");
            return;
        }

        var template = DataCollection.ExtractTableAtPlaceholder(docxBody, tblCellMatchText);
        if (template.FoundTable == null || template.Parent == null || template.Index < 0)
            return;

        var tablesToInsert = new List<Table>();

        foreach (var item in dataList)
        {
            // Clone the template table
            var newTable = (Table)template.FoundTable.CloneNode(true);

            // Replace template placeholders in the cloned table
            foreach (var row in newTable.Elements<TableRow>())
            {
                foreach (var cell in row.Elements<TableCell>())
                {
                    TemplateReplacer.ReplacePlaceholderInTableCell(cell, item, PlaceholderYamlPropertyMap.TableToProperty);
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

    public static void MergeDataSet(Body docxBody, string yamlFilePath)
    {
        List<YamlEntry> dataSet = DataCollection.DeserializeYAML<YamlEntry>(yamlFilePath);
        ReplaceWordTemplate(docxBody, dataSet);
    }

}
