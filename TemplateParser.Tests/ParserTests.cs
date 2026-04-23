using TemplateParser.Core;

namespace TemplateParser.Tests;

public sealed class ParserTests
{
    [Theory]
    [InlineData("test1.docx", "test1.json")]
    [InlineData("test2.docx", "test2.json")]
    public void Cli_Output_Matches_Expected(string docxFile, string expectedJsonFile)
    {
        // Arrange

        // Find the solution root by searching for TemplateParser.sln upward from the test assembly directory
        string FindSolutionRoot()
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null && !File.Exists(Path.Combine(dir.FullName, "TemplateParser.sln")))
            {
                dir = dir.Parent;
            }
            if (dir == null) throw new Exception("Could not find solution root (TemplateParser.sln)");
            return dir.FullName;
        }

        var solutionRoot = FindSolutionRoot();

        string GetSolutionRelativePath(string relativePath)
        {
            return Path.Combine(solutionRoot, relativePath.Replace("/", Path.DirectorySeparatorChar.ToString()));
        }

        var cliProject = GetSolutionRelativePath("TemplateParser.Cli/TemplateParser.Cli.csproj");
        var sampleDocx = GetSolutionRelativePath($"sample-documents/{docxFile}");
        var expectedJsonPath = GetSolutionRelativePath($"expected/{expectedJsonFile}");
        var outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.json");
        var templateId = "11111111-1111-1111-1111-111111111111";

        // Clean up any previous output
        if (File.Exists(outputPath)) File.Delete(outputPath);

        // Act: Run the CLI tool
        var process = new System.Diagnostics.Process();
        process.StartInfo.FileName = "dotnet";
        process.StartInfo.Arguments = $"run --project \"{cliProject}\" -- parse \"{sampleDocx}\" {templateId}";
        process.StartInfo.RedirectStandardOutput = true;
        process.StartInfo.RedirectStandardError = true;
        process.StartInfo.UseShellExecute = false;
        process.Start();
        string stdOut = process.StandardOutput.ReadToEnd();
        string stdErr = process.StandardError.ReadToEnd();
        process.WaitForExit();

        // Assert: CLI ran successfully
        Assert.True(File.Exists(outputPath), $"CLI did not produce {outputPath}. StdErr: {stdErr}");

        // Compare output.json to expected, ignoring 'id' and 'templateId' fields
        var actualJson = File.ReadAllText(outputPath).Trim();
        var expectedJson = File.ReadAllText(expectedJsonPath).Trim();

        using var actualDoc = System.Text.Json.JsonDocument.Parse(actualJson);
        using var expectedDoc = System.Text.Json.JsonDocument.Parse(expectedJson);

        var actualNodes = actualDoc.RootElement.GetProperty("nodes");
        var expectedNodes = expectedDoc.RootElement.GetProperty("nodes");
        Assert.Equal(expectedNodes.GetArrayLength(), actualNodes.GetArrayLength());
        for (int i = 0; i < expectedNodes.GetArrayLength(); i++)
        {
            var expectedNode = expectedNodes[i];
            var actualNode = actualNodes[i];
            // Compare all properties except 'id' and 'templateId'
            foreach (var prop in expectedNode.EnumerateObject())
            {
                if (prop.Name == "id" || prop.Name == "templateId") continue;
                Assert.True(actualNode.TryGetProperty(prop.Name, out var actualProp), $"Missing property '{prop.Name}' in actual node");
                Assert.Equal(prop.Value.ToString(), actualProp.ToString());
            }
        }
    }
}
