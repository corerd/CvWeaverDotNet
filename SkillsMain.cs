using DocumentFormat.OpenXml.Wordprocessing;

public class SkillsMain
{
    private class YamlEntry
    {
        public string? Skill { get; set; }
    }

    private static class PlaceholderYamlPropertyMap
    {
        public const string PlaceholderTableId = "{{tplSkill}}";
        public const string PlaceholderTplDescription = "{{tplSkillList}}";

        // Maps property names of YamlEntry to their corresponding placeholders in the Word table
        public static readonly Dictionary<string, string> PropertyToPlaceholder = new Dictionary<string, string>
        {
            { nameof(YamlEntry.Skill), PlaceholderTplDescription }
        };

        // Maps placeholders back to property names of YamlEntry
        public static readonly Dictionary<string, string> PlaceholderToProperty = new Dictionary<string, string>
        {
            { PlaceholderTableId, "Main" },  // Special: map to fixed string
            { PlaceholderTplDescription, nameof(YamlEntry.Skill) }
        };
    }

    private static void ReplaceWordTemplate(Body docxBody, List<YamlEntry> dataList)
    {
        var template = DataCollection.ExtractTableAtPlaceholder(docxBody, PlaceholderYamlPropertyMap.PlaceholderTableId);
        if (template.FoundTable == null || template.Parent == null || template.Index < 0)
            return;

        // Clone the found table
        var newTable = (Table)template.FoundTable.CloneNode(true);

        // Assuming newTable is a single-row, two-column table
        // get the first (and only) row
        TableRow? templateRow = newTable.Elements<TableRow>().FirstOrDefault();
        if (templateRow == null)
        {
            Console.WriteLine("Cloned table does not contain a row.");
            return;
        }

        // Get the cells (TableCell) within that row
        var cells = templateRow.Elements<TableCell>().ToList();
        if (cells.Count < 2)
        {
            // It is not a two-column table as expected
            Console.WriteLine("Cloned table does not have at least two columns.");
            return;
        }

        // Get and clone the first paragraph of cells[1]
        Paragraph[] cellParagraphs = cells[1].Elements<Paragraph>().ToArray();
        if (cellParagraphs.Length < 1)
        {
            Console.WriteLine("Cloned table cell does not have at least one paragraph.");
            return;
        }
        Paragraph templateParagraph = (Paragraph)cellParagraphs[0].CloneNode(true);

        // Replace template placeholder in cells[0]
        // Since dataItem is null, explicitly tell the method that its type is <YamlEntry>
        TemplateReplacer.ReplacePlaceholderInTableCell<YamlEntry>(
            cells[0],
            null,  // Pass dataItem = null
            PlaceholderYamlPropertyMap.PlaceholderToProperty
        );

        // Update the content of cells[1]
        bool firstItem = true;
        string? propertyNameFound = null;
        foreach (var item in dataList)
        {
            if (firstItem)
            {
                // Replace template placeholder in the first paragraph
                firstItem = false;
                propertyNameFound = TemplateReplacer.ReplacePlaceholderInTableCell(
                    cells[1],
                    item,
                    PlaceholderYamlPropertyMap.PlaceholderToProperty);
                if (propertyNameFound == null)
                    break;
            }
            else
            {
                // Append new paragraphs replacing template placeholder
                Paragraph newParagraph = DataCollection.ReplaceTextInParagraph(templateParagraph, propertyNameFound, item);
                cells[1].AppendChild(newParagraph);
            }
        }

        template.Parent.InsertAt(newTable, template.Index);
    }

    public static void MergeDataSet(Body docxBody, string yamlFilePath)
    {
        List<YamlEntry> dataSet = DataCollection.DeserializeYAML<YamlEntry>(yamlFilePath);
        ReplaceWordTemplate(docxBody, dataSet);
    }

}
