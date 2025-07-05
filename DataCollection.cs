/**
 * Read and parse YAML (.yml) files in C# using a YAML parsing library such as YamlDotNet.
 * https://github.com/aaubry/YamlDotNet
 *
 * Install the package:
 *  dotnet add package YamlDotNet
 */
using System.Collections.Generic;
using System.IO;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

public class DataCollection
{
    /// <summary>
    /// Reads and deserializes a YAML file into a list of objects of a specified generic type.
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
