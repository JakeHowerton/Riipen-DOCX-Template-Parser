using System.Text.Json;
using TemplateParser.Core;

// Status codes for better CLI integration
const int Success = 0;
const int Error = 1;

const string usage = "Usage: dotnet run -- parse <filePath> <templateId> [--out <outputPath>]";

if (args.Length < 3)
{
    Console.Error.WriteLine(usage);
    return Error;
}

var command = args[0];
var filePath = args[1];
var templateIdArg = args[2];

// 1. Validation: Command check
if (!string.Equals(command, "parse", StringComparison.OrdinalIgnoreCase))
{
    Console.Error.WriteLine($"Unsupported command '{command}'. Only 'parse' is supported.");
    return Error;
}

// 2. Validation: File Existence
if (!File.Exists(filePath))
{
    Console.Error.WriteLine($"Error: File not found at {filePath}");
    return Error;
}

// 3. Validation: GUID format
if (!Guid.TryParse(templateIdArg, out var templateId))
{
    Console.Error.WriteLine($"Error: Invalid templateId GUID: {templateIdArg}");
    return Error;
}

var parser = new DocxParser();

try
{
    var result = parser.ParseDocxTemplate(filePath, templateId);

    // Serialization setup: CamelCase and Indented as per contract
    var options = new JsonSerializerOptions
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    };

    var json = JsonSerializer.Serialize(result, options);

    // 4. Output Logic: Check for optional --out flag
    if (args.Length >= 5 && string.Equals(args[3], "--out", StringComparison.OrdinalIgnoreCase))
    {
        var outputPath = args[4];
        File.WriteAllText(outputPath, json);
        Console.WriteLine($"Successfully exported results to: {outputPath}");
    }
    else
    {
        // Default to Standard Output
        Console.WriteLine(json);
    }

    return Success;
}
catch (Exception ex)
{
    Console.Error.WriteLine("An error occurred during parsing:");
    Console.Error.WriteLine(ex.Message);
    return Error;
}