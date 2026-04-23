using System.Text.Json;
using TemplateParser.Core;

const string usage = "Usage: dotnet run -- parse <filePath> <templateId>";

if (args.Length < 3)
{
    Console.Error.WriteLine(usage);
    return;
}

var command = args[0];
var filePath = args[1];
var templateIdArg = args[2];

if (!string.Equals(command, "parse", StringComparison.OrdinalIgnoreCase))
{
    Console.Error.WriteLine("Unsupported command. Only 'parse' is currently available.");
    Console.Error.WriteLine(usage);
    return;
}

if (!File.Exists(filePath))
{
    Console.Error.WriteLine($"File not found: {filePath}");
    return;
}

if (!Guid.TryParse(templateIdArg, out var templateId))
{
    Console.Error.WriteLine($"Invalid templateId GUID: {templateIdArg}");
    return;
}

var parser = new DocxParser();

try
{
    var result = parser.ParseDocxTemplate(filePath, templateId);

    var options = new JsonSerializerOptions
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };
    options.Converters.Add(new NodeJsonConverter());

    var json = JsonSerializer.Serialize(result, options);

    var outputPath = "output.json";
    File.WriteAllText(outputPath, json);
    Console.WriteLine($"Wrote parser output to {outputPath}");
    {
        Console.WriteLine(json);
    }
}
catch (NotImplementedException)
{
    // TODO (Week 6): Replace this temporary message with robust error handling.
    // Example: map known parser exceptions to user-friendly console output and exit codes.
    Console.Error.WriteLine("Parser implementation is intentionally incomplete in this starter repository.");
}
