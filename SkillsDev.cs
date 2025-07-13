using DocumentFormat.OpenXml.Wordprocessing;

public class SkillsDev
{
    private class YamlEntry
    {
        public string? Skill { get; set; }
        public string? Desc { get; set; }
    }

    private static class PlaceholderYamlPropertyMap
    {
        public const string PlaceholderTableId = "{{tplSkillsDev}}";
        public const string PlaceholderTplId = "{{tplSkill}}";
        public const string PlaceholderTplDescription = "{{tplSkillDesc}}";

        // Maps property names of YamlEntry to their corresponding placeholders in the Word table
        public static readonly Dictionary<string, string> PropertyToTable = new Dictionary<string, string>
        {
            { nameof(YamlEntry.Skill), PlaceholderTplId },
            { nameof(YamlEntry.Desc), PlaceholderTplDescription }
        };

        // Maps placeholders back to property names of YamlEntry
        public static readonly Dictionary<string, string> TableToProperty = new Dictionary<string, string>
        {
            { PlaceholderTableId, "Technical aptitude" },  // Special: map to fixed string
            { PlaceholderTplId, nameof(YamlEntry.Skill) },
            { PlaceholderTplDescription, nameof(YamlEntry.Desc) }
        };
    }

    private static void ReplaceWordTemplate(Body docxBody, List<YamlEntry> dataList)
    {
        var template = DataCollection.ExtractTableAtPlaceholder(docxBody, PlaceholderYamlPropertyMap.PlaceholderTableId);
        if (template.FoundTable == null || template.Parent == null || template.Index < 0)
            return;

        var tablesToInsert = new List<Table>();

        bool isFirstCell = true;
        foreach (var item in dataList)
        {
            // Clone the template table
            var newTable = (Table)template.FoundTable.CloneNode(true);

            // Replace template placeholders in the cloned table
            foreach (var row in newTable.Elements<TableRow>())
            {
                foreach (var cell in row.Elements<TableCell>())
                {
                    if (isFirstCell)
                    {
                        // Since dataItem is null, explicitly tell the method that its type is <YamlEntry>
                        TemplateReplacer.ReplacePlaceholderInTableCell<YamlEntry>(
                            cell,
                            null,  // Pass dataItem = null
                            PlaceholderYamlPropertyMap.TableToProperty
                        );
                        isFirstCell = false;
                    }
                    else
                    {
                        TemplateReplacer.ReplacePlaceholderInTableCell(
                            cell,
                            item,
                            PlaceholderYamlPropertyMap.TableToProperty);
                    }
                }
            }
            tablesToInsert.Add(newTable);
        }

        // Insert all new tables at the original index
        int insertIndex = template.Index;
        foreach (var tbl in tablesToInsert)
            template.Parent.InsertAt(tbl, insertIndex++);
    }

    public static void MergeDataSet(Body docxBody, string yamlFilePath)
    {
        List<YamlEntry> dataSet = DataCollection.DeserializeYAML<YamlEntry>(yamlFilePath);
        ReplaceWordTemplate(docxBody, dataSet);
    }

}
