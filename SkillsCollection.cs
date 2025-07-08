using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

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
        { PlaceholderAptitude, nameof(TechAptitudeEntry.Skill) },
        { PlaceholderDescription, nameof(TechAptitudeEntry.Desc) }
    };

    // The PlaceholderTechAptitude placeholder is unique.
    // It is replaced with a fixed substitute string only in the first generated table.
    public const string PlaceholderTechAptitudeSubstitute = "Technical aptitude";
}

public class SkillsCollection
{
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
        ReplacePlaceholderInCell(cells[0], null);

        // Update the content of cells[1]
        bool firstItem = true;
        string? propertyNameFound = null;
        foreach (var item in ApplicationFieldItems)
        {
            if (firstItem)
            {
                // Replace template placeholder in the first paragraph
                firstItem = false;
                propertyNameFound = ReplacePlaceholderInCell(cells[1], item);
                if (propertyNameFound == null)
                    break;
            }
            else
            {
                // Append new paragraphs replacing template placeholder
                Paragraph newParagraph = ReplaceTextInParagraph(templateParagraph, propertyNameFound, item);
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

        bool isFirstTable = true;
        foreach (var item in TechAptitudeItems)
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
                    // var cellText = string.Concat(textElements.Select(t => t.Text));
                    if (textElements.Count < 3) continue;

                    // Look for the pattern: "*{{", someString, "}}"
                    for (int i = 0; i < textElements.Count - 2; i++)
                    {
                        string textStart = textElements[i].Text;
                        bool isStart = textStart.Length >= 2 && textStart.Substring(textStart.Length - 2) == "{{";
                        if (isStart && textElements[i + 2].Text == "}}")
                        {
                            string placeholder = "{{" + textElements[i + 1].Text + "}}";
                            string replaceText = "";
                            if (placeholder == tblCellMatchText)
                            {
                                // tblCellMatchText placeholder is special
                                // Only replace it with some string in the first table
                                // leaving it empty subsequently
                                if (isFirstTable)
                                    replaceText = TechAptitudeEntryPlaceholderMap.PlaceholderTechAptitudeSubstitute;
                            }
                            else
                            {
                                // Map the placeholder to the property name and get its value
                                if (TechAptitudeEntryPlaceholderMap.TableToProperty.TryGetValue(placeholder, out string? propertyName))
                                {
                                    var property = typeof(TechAptitudeEntry).GetProperty(propertyName);
                                    if (property != null)
                                        replaceText = property.GetValue(item)?.ToString() ?? "";
                                }
                            }
                            // Replace the three tokens with the replaceText, preserving formatting
                            // textElements[i].Text is "*{{", then remove the last two characters
                            textElements[i].Text = textElements[i].Text.Substring(0, textElements[i].Text.Length - 2);
                            textElements[i + 1].Text = replaceText;
                            // Assign "" to the last two characters of textElements[i + 2].Text
                            textElements[i + 2].Text = "";
                        }
                    }
                }
            }
            isFirstTable = false;
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

    private static Paragraph ReplaceTextInParagraph(Paragraph srcParagraph, string? propertyName, ApplicationFieldEntry? listItem)
    {
        Paragraph paragraph = (Paragraph)srcParagraph.CloneNode(true);

        // Concatenate all text from the runs in the paragraph
        StringBuilder paragraphTextBuilder = new StringBuilder();
        List<Text> textElements = new List<Text>();
        foreach (Run run in paragraph.Elements<Run>())
        {
            foreach (Text text in run.Elements<Text>())
            {
                paragraphTextBuilder.Append(text.Text);
                textElements.Add(text);
            }
        }
        string currentParagraphText = paragraphTextBuilder.ToString();

        // Get the placeholder by means of he regex search pattern:
        // \{\{     - Matches the literal "{{". The '{' characters are escaped with '\' because they have special meaning in regex.
        // .*?      - Matches any character (.), zero or more times (*), non-greedily (?).
        //            The '?' after '*' makes it non-greedy, meaning it matches the shortest possible string.
        // \}\}     - Matches the literal "}}". The '}' characters are escaped.
        string placeholderPattern = @"\{\{.*?\}\}";
        Match match = Regex.Match(currentParagraphText, placeholderPattern);
        if (!match.Success)
        {
            return paragraph;
        }
        string placeholder = match.Value;
        string replaceText = "";

        // Get the value of the property name
        var property = typeof(ApplicationFieldEntry).GetProperty(propertyName);
        if (property != null)
            replaceText = property.GetValue(listItem)?.ToString() ?? "";

        // Proceed with replacement
        // This is a simplified approach that assumes we can find and replace.
        // For complex scenarios (multiple replacements in one paragraph, partial replacements),
        // a more sophisticated algorithm is needed (e.g., breaking down runs into single characters,
        // then reassembling).
        // However, for typical full-string replacements, this can work.

        // Get the RunProperties of the first run in the paragraph.
        // This is a simplification. For truly preserving formatting across multiple runs,
        // you'd need to identify the exact run properties for the *start* of the match.
        RunProperties? firstRunProperties = paragraph.Elements<Run>().FirstOrDefault()?.RunProperties;

        // Remove all existing runs from the paragraph
        // WARNING: This will remove ALL formatting from the original runs.
        // A more robust solution involves carefully manipulating runs around the replacement.
        paragraph.RemoveAllChildren<Run>();

        // Create a new Run for the replaced text
        Run newRun = new Run();

        // // Add the text to the Run
        Text newText = new Text(currentParagraphText.Replace(placeholder, replaceText));
        // Preserve the formatting of the first run (if available)
        if (firstRunProperties != null)
        {
            newRun.AppendChild((RunProperties)firstRunProperties.CloneNode(true));
        }
        // Preserve spaces if they are significant
        if (newText.Text.Contains(" ") || newText.Text.Contains("\t"))
        {
            newText.Space = SpaceProcessingModeValues.Preserve;
        }
        newRun.AppendChild(newText);

        // Add the Run to the Paragrap
        paragraph.AppendChild(newRun);

        // TODO: If you need to handle multiple occurrences within the same paragraph
        // or preserve formatting for parts of the paragraph *before* and *after* the replacement,
        // you'll need to implement a more advanced algorithm as described in Eric White's blog
        // (see search results). This typically involves breaking runs into individual characters.

        return paragraph;
    }

    private static string? ReplacePlaceholderInCell(TableCell cell, ApplicationFieldEntry? listItem)
    {
        // Gather all Text elements in this cell
        var textElements = cell.Descendants<Text>().ToList();
        if (textElements.Count < 3)
            return null;

        // Look for the pattern: "*{{", someString, "}}"
        string? propertyName = null;
        for (int i = 0; i < textElements.Count - 2; i++)
        {
            string textStart = textElements[i].Text;
            bool isStart = textStart.Length >= 2 && textStart.Substring(textStart.Length - 2) == "{{";
            if (isStart && textElements[i + 2].Text == "}}")
            {
                string placeholder = "{{" + textElements[i + 1].Text + "}}";
                string replaceText = "";
                // Map the placeholder to the property name and get its value
                if (ApplicationFieldEntryPlaceholderMap.TableToProperty.TryGetValue(placeholder, out propertyName))
                {
                    // placeholder found
                    if (listItem == null)
                        replaceText = propertyName;
                    else
                    {
                        var property = typeof(ApplicationFieldEntry).GetProperty(propertyName);
                        if (property != null)
                            replaceText = property.GetValue(listItem)?.ToString() ?? "";
                    }
                }
                else
                {
                    // placeholder not found
                    propertyName = null;
                }

                // Replace the three tokens with the replaceText, preserving formatting
                // textElements[i].Text is "*{{", then remove the last two characters
                textElements[i].Text = textElements[i].Text.Substring(0, textElements[i].Text.Length - 2);
                textElements[i + 1].Text = replaceText;
                // Assign "" to the last two characters of textElements[i + 2].Text
                textElements[i + 2].Text = "";
            }
        }
        return propertyName;
    }
}
