using DocumentFormat.OpenXml.Wordprocessing;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

public static class TemplateReplacer
{
    /// <summary>
    /// Replaces text patterns like "*{{placeholder}}*" in a single MS Word table cell.
    /// Word documents often split text across multiple Run and Text elements due to formatting, spellcheck, or editing.
    /// The text "Hello {{someString}} World" might be fragmented like:
    /// Run1: "Hello {{so"
    /// Run2: "meStr"
    /// Run3: "ing}} World"
    /// This method can handle patterns that span across multiple Run/Text elements.
    /// </summary>
    /// <typeparam name="T">The type of the data item to retrieve replacement values from.</typeparam>
    /// <param name="cell">The TableCell to inspect and modify.</param>
    /// <param name="dataItem">The object containing data to replace placeholders with. Can be null.</param>
    /// <param name="tableToPropertyMap">A dictionary mapping placeholder strings (e.g., "{{placeholder}}") to
    /// property names of T or fixed strings.</param>
    /// <returns>The property name or fixed string that was replaced, or null if no match was found.</returns>
    public static string? ReplacePlaceholderInTableCell<T>(
        TableCell cell,
        T? dataItem,
        Dictionary<string, string> tableToPropertyMap) where T : class
    {
        var texts = cell.Descendants<Text>().ToList();
        var fullText = string.Join("", texts.Select(t => t.Text));

        // Regex to find {{someString}} pattern
        var regex = new Regex(@"(\{\{.*?\}\})");
        var matches = regex.Matches(fullText);

        if (matches.Count == 0)
        {
            return null; // No pattern found in this cell
        }

        string? replacedPropertyName = null;

        foreach (Match match in matches)
        {
            string placeholder = match.Value;
            var found = SearchPlaceholderProperty(placeholder, dataItem, tableToPropertyMap);
            if (found.PropertyName == null || found.PropertyValue == null)
            {
                // Placeholder and/or propertyName not found
                // then skip and move to next match
                continue;
            }

            replacedPropertyName = found.PropertyName;
            string newText = placeholder.Replace(match.Groups[1].Value, found.PropertyValue);

            // Replace only the placeholder text across Text elements
            ReplaceAcrossRuns(texts, match.Index, match.Length, newText);
        }

        return replacedPropertyName;
    }

    private static (string? PropertyName, string? PropertyValue) SearchPlaceholderProperty<T>(
        string placeholder,
        T? dataItem,
        Dictionary<string, string> tableToPropertyMap) where T : class
    {
        string? propertyValue = null;

        // Map the placeholder to its replacement value
        if (tableToPropertyMap.TryGetValue(placeholder, out string? propertyName))
        {
            if (dataItem == null)
            {
                // Use the property name/fixed string itself
                propertyValue = propertyName;
            }
            else
            {
                // Handle the actual properties
                PropertyInfo? property = typeof(T).GetProperty(propertyName);
                if (property != null)
                {
                    if (property.PropertyType == typeof(List<string>))
                    {
                        List<string>? fieldsList = property.GetValue(dataItem) as List<string>;
                        propertyValue = string.Join(", ", fieldsList ?? new List<string>());
                    }
                    else
                        propertyValue = property.GetValue(dataItem)?.ToString() ?? "";
                }
                else
                {
                    // Property '{propertyName}' not found on type '{typeof(T).Name}' for placeholder '{placeholder}'
                    // then return '{propertyValue}' as an empty string.
                    propertyValue = "";
                }
            }
        }
        // else placeholder not found in map

        return (propertyName, propertyValue);
    }

    private static void ReplaceAcrossRuns(List<Text> texts, int startIndex, int length, string newText)
    {
        int currentIndex = 0;
        int replaced = 0;

        foreach (var text in texts)
        {
            var chars = text.Text.ToCharArray();
            var sb = new StringBuilder();

            for (int i = 0; i < chars.Length; i++)
            {
                if (currentIndex >= startIndex && replaced < length)
                {
                    if (replaced == 0)
                        sb.Append(newText);
                    replaced++;
                }
                else
                {
                    sb.Append(chars[i]);
                }
                currentIndex++;
            }

            text.Text = sb.ToString();
            if (replaced >= length) break;
        }
    }
}
