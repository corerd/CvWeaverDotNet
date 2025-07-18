﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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

        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputDocName, true))
        {
            if (wordDoc.MainDocumentPart == null || wordDoc.MainDocumentPart.Document == null)
            {
                Console.WriteLine($"The document does not contain a main document part or document.");
                return;
            }

            // Get a reference to the main content area (the 'Body') of the Word document.
            // This 'body' object will be passed as an argument to methods, allowing them
            // to locate and replace/update specific data within the document's content.
            Body? body = wordDoc.MainDocumentPart.Document.Body;
            if (body == null)
            {
                Console.WriteLine("The document body is null.");
                return;
            }

            Console.WriteLine($"Get '{DataStore.ApplicationFieldPath}'");
            List<DataCollection.Domain> applicationFields =
                DataCollection.DeserializeYAML<DataCollection.Domain>(DataStore.ApplicationFieldPath);

            Console.WriteLine($"Merge '{DataStore.SkillsPath}'");
            SkillsMain.MergeDataSet(body, DataStore.SkillsPath);

            Console.WriteLine($"Merge '{DataStore.ApplicationFieldPath}'");
            ApplicationFields.MergeDataSet(body, applicationFields);

            Console.WriteLine($"Merge '{DataStore.SkillsDevPath}'");
            SkillsDev.MergeDataSet(body, DataStore.SkillsDevPath);

            Console.WriteLine($"Merge '{DataStore.HistoryPath}' with:");
            Console.WriteLine($"      '{DataStore.ApplicationFieldPath}'");
            Console.WriteLine($"      '{DataStore.HyperlinkDescPath}'");
            History.MergeDataSet(body, DataStore.HistoryPath, DataStore.HyperlinkDescPath, applicationFields);

            // Save both the document and its main part
            wordDoc.MainDocumentPart.Document.Save();
            wordDoc.Save();
        }

        Console.WriteLine($"Output '{outputDocName}'");
        Console.WriteLine("Done!");
    }
}
