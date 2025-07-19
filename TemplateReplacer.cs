using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

public static class TemplateReplacer
{
    public static void ReplaceHyperlinkTag(
        Body ancestorBody,
        List<Paragraph> paragraphList,
        Dictionary<string, DataCollection.HyperlinkDesc> hyperlinkDictionary)
    {
        // Define a tag pattern as ((abbreviation))
        const string tagPattern = @"\(\([^)]+\)\)";

        // Add hyperlink relationship to the main document part
        var mainPart = ancestorBody?.Ancestors<Document>().FirstOrDefault()?.MainDocumentPart;
        if (mainPart == null)
            return;  // no Main Part

        foreach (var para in paragraphList)
        {
            var runs = para.Descendants<Run>()
                //.Where(r => r.InnerText.Contains(textToReplace))
                .Where(r => Regex.IsMatch(r.InnerText, tagPattern))
                .ToList();

            foreach (var run in runs)
            {
                var runText = run.GetFirstChild<Text>();
                if (runText == null) continue;

                string fullText = runText.Text;

                // Find all matches of tagPattern in the run's text
                var matches = Regex.Matches(fullText, "(" + tagPattern + ")");
                if (matches.Count == 0) continue;

                int lastIndex = 0;
                List<OpenXmlElement> newElements = new List<OpenXmlElement>();

                foreach (Match match in matches)
                {
                    string tagKey = match.Groups[1].Value;
                    if (!hyperlinkDictionary.TryGetValue(tagKey, out var hyperlinkDesc))
                        continue;

                    string hyperlinkText = hyperlinkDesc.Name ?? string.Empty;
                    string hyperlinkUrl = hyperlinkDesc.Link ?? string.Empty;

                    // Add text before the match
                    if (match.Index > lastIndex)
                    {
                        string beforeText = fullText.Substring(lastIndex, match.Index - lastIndex);
                        if (!string.IsNullOrEmpty(beforeText))
                        {
                            Run beforeRun = run.RunProperties != null
                                ? new Run(run.RunProperties.CloneNode(true))
                                : new Run();
                            beforeRun.AppendChild(new Text(beforeText) { Space = SpaceProcessingModeValues.Preserve });
                            newElements.Add(beforeRun);
                        }
                    }

                    // Ensure the URL is absolute; if not, prepend the base URL
                    Uri? hyperlinkUri;
                    if (!Uri.TryCreate(hyperlinkUrl, UriKind.Absolute, out hyperlinkUri))
                    {
                        // URL is not absolute
                        // then combine with base URL
                        hyperlinkUri = new Uri(new Uri(DataStore.HyperlinkBaseUrl), hyperlinkUrl);
                    }

                    // Add hyperlink relationship
                    var hyperlinkRel = mainPart.AddHyperlinkRelationship(
                        hyperlinkUri,
                        true);

                    // Create new RunProperties for hyperlink with custom style
                    RunProperties hyperlinkRunProperties = new RunProperties();

                    // Set custom color (e.g., blue)
                    Color color = new Color() { Val = "0000FF" };  // Hex RGB for blue
                    hyperlinkRunProperties.Append(color);

                    // Set underline (single)
                    Underline underline = new Underline() { Val = UnderlineValues.Single };
                    hyperlinkRunProperties.Append(underline);

                    // Optionally set font size (half-point, e.g., 24=12pt)
                    // FontSize fontSize = new FontSize() { Val = "24" };
                    // hyperlinkRunProperties.Append(fontSize);

                    // Create hyperlink element
                    Hyperlink hyperlink = new Hyperlink()
                    {
                        History = OnOffValue.FromBoolean(true),
                        Id = hyperlinkRel.Id
                    };

                    // Create the run inside the hyperlink with custom properties and text
                    Run hyperlinkTextRun = new Run();
                    hyperlinkTextRun.Append(hyperlinkRunProperties);
                    hyperlinkTextRun.Append(new Text(hyperlinkText) { Space = SpaceProcessingModeValues.Preserve });

                    hyperlink.AppendChild(hyperlinkTextRun);
                    newElements.Add(hyperlink);

                    lastIndex = match.Index + match.Length;
                }

                // Add text after the last match
                if (lastIndex < fullText.Length)
                {
                    string afterText = fullText.Substring(lastIndex);
                    if (!string.IsNullOrEmpty(afterText))
                    {
                        Run afterRun = run.RunProperties != null
                            ? new Run(run.RunProperties.CloneNode(true))
                            : new Run();
                        afterRun.AppendChild(new Text(afterText) { Space = SpaceProcessingModeValues.Preserve });
                        newElements.Add(afterRun);
                    }
                }

                // Replace the original run with the new elements
                if (run.Parent != null && newElements.Count > 0)
                {
                    run.Parent.InsertBefore(newElements.First(), run);
                    foreach (var elem in newElements.Skip(1))
                    {
                        run.Parent.InsertAfter(elem, newElements.First());
                        newElements[0] = elem;
                    }
                    run.Remove();
                }
            }

        }

    }

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
    /// <param name="domainDictionary">Optional dictionary mapping domain Id to Name.</param>
    /// <returns>The property name or fixed string that was replaced, or null if no match was found.</returns>
    public static string? ReplacePlaceholderInTableCell<T>(
        TableCell cell,
        T? dataItem,
        Dictionary<string, string> tableToPropertyMap,
        Dictionary<string, DataCollection.Domain>? domainDictionary = null) where T : class
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
            var found = SearchPlaceholderProperty(placeholder, dataItem, tableToPropertyMap, domainDictionary);
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
        Dictionary<string, string> tableToPropertyMap,
        Dictionary<string, DataCollection.Domain>? domainDictionary) where T : class
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
                        List<string>? domainIdList = property.GetValue(dataItem) as List<string>;
                        if (domainIdList != null)
                        {
                            if (domainDictionary != null)
                                propertyValue = JoinDomains(domainIdList, domainDictionary);
                            else
                                propertyValue = string.Join(", ", domainIdList.Select(item => $"{item}"));
                        }
                        else
                            propertyValue = "";
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

    private static string JoinDomains(
        List<string> domainIdList,
        Dictionary<string, DataCollection.Domain> domainDictionary)
    {
        var resultParts = new List<string>();

        foreach (string id in domainIdList)
        {
            if (domainDictionary.TryGetValue(id, out DataCollection.Domain? domain))
            {
                // ID found in dictionary, use the domain's Name (if not null)
                if (domain.Name != null)
                {
                    resultParts.Add(domain.Name);
                }
                else
                {
                    // If domain.Name is null, append the ID itself
                    resultParts.Add(id);
                }
            }
            else
            {
                // ID not found in dictionary, append the ID itself
                resultParts.Add(id);
            }
        }

        return string.Join(", ", resultParts);
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
