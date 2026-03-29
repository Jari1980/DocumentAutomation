using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;

class WordService
{
    private readonly string _templatePath;
    private readonly string _outputFolder;

    public WordService(string templatePath, string outputFolder)
    {
        _templatePath = templatePath;
        _outputFolder = outputFolder;
        Directory.CreateDirectory(_outputFolder);
    }

    public void Generate(Dictionary<string, string> replacements, string outputFileName)
    {
        string tempFolder = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(tempFolder);

        ZipFile.ExtractToDirectory(_templatePath, tempFolder);

        string contentXml = Path.Combine(tempFolder, "content.xml");
        string xmlText = File.ReadAllText(contentXml);

        foreach (var kv in replacements)
        {
            xmlText = xmlText.Replace(kv.Key, kv.Value);
        }

        File.WriteAllText(contentXml, xmlText);

        string outputFile = Path.Combine(_outputFolder, outputFileName);
        if (File.Exists(outputFile)) File.Delete(outputFile);
        ZipFile.CreateFromDirectory(tempFolder, outputFile);

        Directory.Delete(tempFolder, true);
        Console.WriteLine("Generated: " + outputFile);
    }
}