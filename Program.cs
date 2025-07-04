public class CvWeaverDotNet
{   
    public static void Main(string[] args)
    {
        string initialDocName = DataStore.DocTemplatePath;
        string directory = Path.GetDirectoryName(initialDocName) ?? string.Empty;
        string outputDocName = Path.Combine(
            directory,
            Path.GetFileNameWithoutExtension(initialDocName) + "_merged" +
            Path.GetExtension(initialDocName)
        );

        // Copy the initial document to the output document, preserving the original
        File.Copy(initialDocName, outputDocName, true);

        List<HistoryEntry> historyItems = HistoryCollection.Deserialize(DataStore.HistoryPath);
        HistoryCollection.ReplaceHistory(outputDocName, historyItems);

        Console.WriteLine($"Done");
    }
}
