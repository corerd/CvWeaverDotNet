# Weaving CV resume into a MS Word document

Automate the generation of Microsoft Word resumes using C# and .NET
by merging dynamic career information with a predefined DOCX template.

This approach leverages the built-in capabilities of the Open XML SDK,
so no additional third-party libraries are necessary.

## Project Setup

1.  Create a new C\# Console Application:
```bash
dotnet new console -n CvWeaverDotNet
cd CvWeaverDotNet
```

2.  Add the Open XML SDK NuGet package:
```bash
dotnet add package DocumentFormat.OpenXml
```
