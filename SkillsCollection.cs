using DocumentFormat.OpenXml.Wordprocessing;

public class SkillEntry
{
    public string? Skill { get; set; }
}

public static class SkillEntryPlaceholderMap
{
    public const string PlaceholderSkill = "{{tplSkill}}";
    public const string PlaceholderDescription = "{{tplSkillList}}";

    // Maps property names of ApplicationFieldEntry to their corresponding placeholders in the Word table
    public static readonly Dictionary<string, string> PropertyToTable = new Dictionary<string, string>
    {
        { nameof(SkillEntry.Skill), PlaceholderDescription }
    };

    // Maps placeholders back to property names of ApplicationFieldEntry
    public static readonly Dictionary<string, string> TableToProperty = new Dictionary<string, string>
    {
        { PlaceholderSkill, "Main" },  // fixed string
        { PlaceholderDescription, nameof(SkillEntry.Skill) }
    };
}

public class ApplicationFieldEntry
    {
        public string? Id { get; set; }
        public string? Name { get; set; }
        public string? Desc { get; set; }
    }

public static class ApplicationFieldEntryPlaceholderMap
{
    public const string PlaceholderApplicationField = "{{tplAppField}}";
    public const string PlaceholderDescription = "{{tplAppFieldList}}";

    // Maps property names of ApplicationFieldEntry to their corresponding placeholders in the Word table
    public static readonly Dictionary<string, string> PropertyToTable = new Dictionary<string, string>
    {
        { nameof(ApplicationFieldEntry.Desc), PlaceholderDescription }
    };

    // Maps placeholders back to property names of ApplicationFieldEntry
    public static readonly Dictionary<string, string> TableToProperty = new Dictionary<string, string>
    {
        { PlaceholderApplicationField, "Application fields" },  // fixed string
        { PlaceholderDescription, nameof(ApplicationFieldEntry.Desc) }
    };
}

public class TechAptitudeEntry
{
    public string? Skill { get; set; }
    public string? Desc { get; set; }
}

public static class TechAptitudeEntryPlaceholderMap
{
    public const string PlaceholderTechAptitude = "{{tplSkillsDev}}";
    public const string PlaceholderAptitude = "{{tplSkill}}";
    public const string PlaceholderDescription = "{{tplSkillDesc}}";

    // Maps property names of TechAptitudeEntry to their corresponding placeholders in the Word table
    public static readonly Dictionary<string, string> PropertyToTable = new Dictionary<string, string>
    {
        { nameof(TechAptitudeEntry.Skill), PlaceholderAptitude },
        { nameof(TechAptitudeEntry.Desc), PlaceholderDescription }
    };

    // Maps placeholders back to property names of TechAptitudeEntry
    public static readonly Dictionary<string, string> TableToProperty = new Dictionary<string, string>
    {
        { PlaceholderTechAptitude, "Technical aptitude" },  // fixed string
        { PlaceholderAptitude, nameof(TechAptitudeEntry.Skill) },
        { PlaceholderDescription, nameof(TechAptitudeEntry.Desc) }
    };
}

public class SkillsCollection
{
    public static void ReplaceSkillTemplate(Body docxBody, List<SkillEntry> skillItems)
    {
        var template = DataCollection.ExtractTableAtPlaceholder(docxBody, SkillEntryPlaceholderMap.PlaceholderSkill);
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
        // Since listItem is null, explicitly tell the method that T is <ApplicationFieldEntry>
        DataCollection.ReplacePlaceholderInCell<SkillEntry>(
            cells[0],
            null,  // Pass listItem = null
            SkillEntryPlaceholderMap.TableToProperty
        );

        // Update the content of cells[1]
        bool firstItem = true;
        string? propertyNameFound = null;
        foreach (var item in skillItems)
        {
            if (firstItem)
            {
                // Replace template placeholder in the first paragraph
                firstItem = false;
                propertyNameFound = DataCollection.ReplacePlaceholderInCell(
                    cells[1],
                    item,
                    SkillEntryPlaceholderMap.TableToProperty
                );
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

    public static void MergeSkillData(Body docxBody, string dataSetFilePath)
    {
        List<SkillEntry> skillItems = DataCollection.DeserializeYAML<SkillEntry>(dataSetFilePath);
        ReplaceSkillTemplate(docxBody, skillItems);
    }

    public static void ReplaceApplicationFieldTemplate(Body docxBody, List<ApplicationFieldEntry> ApplicationFieldItems)
    {
        var template = DataCollection.ExtractTableAtPlaceholder(docxBody, ApplicationFieldEntryPlaceholderMap.PlaceholderApplicationField);
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
        // Since listItem is null, explicitly tell the method that T is <ApplicationFieldEntry>
        DataCollection.ReplacePlaceholderInCell<ApplicationFieldEntry>(
            cells[0],
            null,  // Pass listItem = null
            ApplicationFieldEntryPlaceholderMap.TableToProperty
        );

        // Update the content of cells[1]
        bool firstItem = true;
        string? propertyNameFound = null;
        foreach (var item in ApplicationFieldItems)
        {
            if (firstItem)
            {
                // Replace template placeholder in the first paragraph
                firstItem = false;
                propertyNameFound = DataCollection.ReplacePlaceholderInCell(
                    cells[1],
                    item,
                    ApplicationFieldEntryPlaceholderMap.TableToProperty
                );
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

    public static void MergeApplicationFieldData(Body docxBody, string dataSetFilePath)
    {
        List<ApplicationFieldEntry> applicationFieldItems = DataCollection.DeserializeYAML<ApplicationFieldEntry>(dataSetFilePath);
        ReplaceApplicationFieldTemplate(docxBody, applicationFieldItems);
    }

    public static void ReplaceTechAptitudeTemplate(Body docxBody, List<TechAptitudeEntry> TechAptitudeItems)
    {
        const string tblCellMatchText = TechAptitudeEntryPlaceholderMap.PlaceholderTechAptitude;

        var template = DataCollection.ExtractTableAtPlaceholder(docxBody, tblCellMatchText);
        if (template.FoundTable == null || template.Parent == null || template.Index < 0)
            return;

        var tablesToInsert = new List<Table>();

        bool isFirstCell = true;
        foreach (var item in TechAptitudeItems)
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
                        // Since listItem is null, explicitly tell the method that T is <TechAptitudeEntry>
                        DataCollection.ReplacePlaceholderInCell<TechAptitudeEntry>(
                            cell,
                            null,  // Pass listItem = null
                            TechAptitudeEntryPlaceholderMap.TableToProperty
                        );
                        isFirstCell = false;
                    }
                    else
                    {
                        DataCollection.ReplacePlaceholderInCell(
                            cell,
                            item,
                            TechAptitudeEntryPlaceholderMap.TableToProperty
                        );
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

    public static void MergeTechAptitudeData(Body docxBody, string dataSetFilePath)
    {
        List<TechAptitudeEntry> techAptitudeItems = DataCollection.DeserializeYAML<TechAptitudeEntry>(dataSetFilePath);
        ReplaceTechAptitudeTemplate(docxBody, techAptitudeItems);
    }
}
