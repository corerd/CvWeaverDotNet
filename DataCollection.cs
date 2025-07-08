using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

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
