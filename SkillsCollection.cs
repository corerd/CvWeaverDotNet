using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
    public static void ReplaceTechAptitudeTemplate(string docFilePath, List<TechAptitudeEntry> TechAptitudeItems)
    {
        const string tblCellMatchText = TechAptitudeEntryPlaceholderMap.PlaceholderTechAptitude;

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

            bool isFirstTable = true;
            foreach (var item in TechAptitudeItems)
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
            int insertIndex = index;
            foreach (var tbl in tablesToInsert)
                parent.InsertAt(tbl, insertIndex++);

            // Save both the document and its main part
            wordDoc.MainDocumentPart.Document.Save();
            wordDoc.Save();
        }
    }

    public static void MergeTechAptitudeData(string docFilePath, string dataSetFilePath)
    {
        List<TechAptitudeEntry> techAptitudeItems = DataCollection.DeserializeYAML<TechAptitudeEntry>(DataStore.SkillDevPath);
        ReplaceTechAptitudeTemplate(docFilePath, techAptitudeItems);
   }
}
