using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

public class DataCollection
{
    public static (Table? FoundTable, OpenXmlElement? Parent, int Index) ExtractTableAtPlaceholder(Body docxBody, string firstCellPlaceholder)
    {
        // Find the original table by matching the content of cell (0,0)
        Table? foundTable = null;
        foreach (var table in docxBody.Elements<Table>())
        {
            var firstRow = table.Elements<TableRow>().FirstOrDefault();
            var firstCell = firstRow?.Elements<TableCell>().FirstOrDefault();
            var cellText = firstCell?.InnerText;

            if (!string.IsNullOrEmpty(cellText) && cellText.Contains(firstCellPlaceholder))
            {
                foundTable = table;
                break;
            }
        }

        if (foundTable == null)
        {
            Console.WriteLine($"No table found with cell (0,0) text matching '{firstCellPlaceholder}'");
            return (null, null, -1);
        }

        // Capture index of original table in its parent
        var parent = foundTable.Parent;
        if (parent == null)
        {
            Console.WriteLine("The template table's parent is null.");
            return (null, null, -1);
        }
        int index = parent.ChildElements.ToList().IndexOf(foundTable);

        // Remove original table from the document
        foundTable.Remove();

        return (foundTable, parent, index);
    }

    /// <summary>
    /// Replaces placeholders in a given TableCell with values from a generic list item.
    /// The method assumes placeholders are structured as: 'partial_text{{', 'PLACEHOLDER_KEY', '}}'.
    /// </summary>
    /// <typeparam name="T">The type of the list item containing the properties to substitute.</typeparam>
    /// <param name="cell">The TableCell to search and replace within.</param>
    /// <param name="listItem">The object instance from which to retrieve property values for substitution.</param>
    /// <param name="tableToPropertyMap">A dictionary mapping placeholder strings to property names (or fixed strings).</param>
    /// <returns>The property name that was replaced, or null if no matching placeholder was found.</returns>
    public static string? ReplacePlaceholderInCell<T>(
        TableCell cell,
        T? listItem,
        Dictionary<string, string> tableToPropertyMap) where T : class // Constraint T to be a class (reference type)
    {
        // Gather all Text elements in this cell. Using Descendants<Text>() is robust.
        var textElements = cell.Descendants<Text>().ToList();

        // If there are fewer than 3 text elements, the pattern "{{PLACEHOLDER}}" cannot exist.
        // It should be at least: Text containing "{{", Text with placeholder name, Text containing "}}"
        if (textElements.Count < 3)
            return null;

        string? replacedPropertyName;

        // Look for the pattern: "partial_text{{", "someString", "}}"
        // Iterate through text elements, checking for the {{ and }} markers
        for (int i = 0; i < textElements.Count - 2; i++)
        {
            string textStart = textElements[i].Text;
            string textEnd = textElements[i + 2].Text;

            // Check if the current Text element ends with "{{"
            bool isStartMarker = textStart.EndsWith("{{");
            // Check if the Text element two positions ahead is exactly "}}"
            bool isEndMarker = textEnd == "}}";

            if (isStartMarker && isEndMarker)
            {
                string placeholderContent = textElements[i + 1].Text;
                string fullPlaceholder = "{{" + placeholderContent + "}}";
                string replaceText = "";

                // Map the full placeholder to the property name (or fixed string)
                if (tableToPropertyMap.TryGetValue(fullPlaceholder, out replacedPropertyName))
                {
                    // Placeholder found in the map
                    if (listItem == null)
                    {
                        // If listItem is null, replace with the property name or fixed string itself
                        replaceText = replacedPropertyName;
                    }
                    else
                    {
                        // Attempt to get the property value from the listItem
                        PropertyInfo? property = typeof(T).GetProperty(replacedPropertyName);
                        if (property != null)
                        {
                            // Get the value and convert to string, handling null values
                            replaceText = property.GetValue(listItem)?.ToString() ?? "";
                        }
                    }

                    // --- Perform the replacement in the Open XML elements ---
                    // textElements[i] contains "partial_text{{"
                    // Remove the last two characters "{{"
                    textElements[i].Text = textElements[i].Text.Substring(0, textElements[i].Text.Length - 2);

                    // textElements[i+1] contains the actual placeholder name, replace with the value
                    textElements[i + 1].Text = replaceText;

                    // textElements[i+2] contains "}}", assign "" to effectively remove it
                    textElements[i + 2].Text = "";

                    // Since we found and replaced a placeholder, we can return.
                    return replacedPropertyName;
                }
            }
        }
        return null; // No matching placeholder found and replaced
    }

    public static Paragraph ReplaceTextInParagraph<T>(
        Paragraph srcParagraph,
        string? propertyName,
        T? listItem) where T : class
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
        if (!string.IsNullOrEmpty(propertyName))
        {
            PropertyInfo? property = typeof(T).GetProperty(propertyName);
            if (property != null)
                replaceText = property.GetValue(listItem)?.ToString() ?? "";
        }

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

    /// <summary>
    /// Reads and deserializes a YAML file into a list of objects of a specified generic type.
    /// Use the YAML parsing library YamlDotNet:
    ///     https://github.com/aaubry/YamlDotNet
    /// Installing the package:
    ///     dotnet add package YamlDotNet
    /// </summary>
    /// <typeparam name="T">The type of objects to deserialize into a list.</typeparam>
    /// <param name="path">The file path to the YAML document.</param>
    /// <returns>A list of objects of type T, deserialized from the YAML, or an empty list if deserialization results in null.</returns>
    public static List<T> DeserializeYAML<T>(string yamlFilePath) where T : class, new()
    {
        var deserializer = new DeserializerBuilder()
            .WithNamingConvention(CamelCaseNamingConvention.Instance)
            .Build();

        using var reader = new StreamReader(yamlFilePath);
        var yaml = reader.ReadToEnd();

        // Deserialize the YAML into a List of type T
        var entries = deserializer.Deserialize<List<T>>(yaml);

        // Return the deserialized list, or a new empty list if it's null
        return entries ?? new List<T>();
    }
}
