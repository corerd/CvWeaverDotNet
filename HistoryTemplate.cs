/**
 * Using Open XML SDK
 *
 * dotnet add package DocumentFormat.OpenXml
 * 
 * Alternative is DocX library by Xceed:
 * dotnet add package Xceed.Words.NET  
 */
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Validation;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml;


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

public static class HistoryTemplate
{
    public static void ReplaceHistory(string filePath, List<HistoryEntry> historyItems)
    {
        //string tblCellMatchText = "{{tbl_history_year}}";
        // Get the table name for Year property
        HistoryEntryTemplateMap.PropertyToTable.TryGetValue(nameof(HistoryEntry.Year), out string? tblCellMatchText);

        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
        {
            var body = wordDoc.MainDocumentPart.Document.Body;

            // Find the original table by matching the content of cell (0,0)
            Table templateTable = null;
            foreach (var table in body.Elements<Table>())
            {
                var firstRow = table.Elements<TableRow>().FirstOrDefault();
                var firstCell = firstRow?.Elements<TableCell>().FirstOrDefault();
                var cellText = firstCell?.InnerText;

                if (!string.IsNullOrEmpty(cellText) && cellText.Contains(tblCellMatchText))
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

    // Append a table next to the parent
    public static int AppendTable(int index, OpenXmlElement parent, Table tableContent)
    {
        // Append a line space (empty paragraph) after the first inserted table
        var emptyParagraph = new Paragraph(new Run(new Text("")));
        parent.InsertAt(emptyParagraph, ++index);

        // Append another instance of TableContent after the line space
        var newTableClone = (Table)tableContent.CloneNode(true);
        parent.InsertAt(newTableClone, ++index);

        return index;
    }

    // Example method to generate a basic new table
    public static Table CreateSampleTable()
    {
        var table = new Table();

        var row1 = new TableRow();
        row1.Append(
            new TableCell(new Paragraph(new Run(new Text("New Cell 1")))),
            new TableCell(new Paragraph(new Run(new Text("New Cell 2"))))
        );

        var row2 = new TableRow();
        row2.Append(
            new TableCell(new Paragraph(new Run(new Text("")))),
            new TableCell(new Paragraph(new Run(new Text("New Cell 4"))))
        );

        table.Append(row1);
        table.Append(row2);

        return table;
    }
}
