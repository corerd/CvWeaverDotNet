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

        Console.WriteLine($"Merge '{DataStore.SkillDevPath}'");
        List<TechAptitudeEntry> TechAptitudeItems = SkillsCollection.DeserializeTechAptitude(DataStore.SkillDevPath);
        SkillsCollection.ReplaceTechAptitude(outputDocName, TechAptitudeItems);

        Console.WriteLine($"Merge '{DataStore.HistoryPath}'");
        List<HistoryEntry> historyItems = HistoryCollection.Deserialize(DataStore.HistoryPath);
        HistoryCollection.ReplaceHistory(outputDocName, historyItems);

        Console.WriteLine($"Output '{outputDocName}'");
        Console.WriteLine("Done!");
    }
}
