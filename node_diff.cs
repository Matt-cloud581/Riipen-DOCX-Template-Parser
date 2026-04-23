// This script compares the node lists in expected/test1.json and TemplateParser.Tests/bin/Debug/net10.0/output.json
// and prints a diff by type, title, and parentId.
using System;
using System.IO;
using System.Text.Json;
using System.Collections.Generic;

string expectedPath = @"expected/test1.json";
string actualPath = @"TemplateParser.Tests/bin/Debug/net10.0/output.json";

JsonDocument expectedDoc = JsonDocument.Parse(File.ReadAllText(expectedPath));
JsonDocument actualDoc = JsonDocument.Parse(File.ReadAllText(actualPath));

var expectedNodes = expectedDoc.RootElement.GetProperty("nodes");
var actualNodes = actualDoc.RootElement.GetProperty("nodes");

var expectedSet = new HashSet<string>();
foreach (var node in expectedNodes.EnumerateArray())
{
    string key = $"{node.GetProperty("type").GetString()}|{node.GetProperty("title").GetString()}|{node.GetProperty("parentId").GetString()}";
    expectedSet.Add(key);
}

var actualSet = new HashSet<string>();
foreach (var node in actualNodes.EnumerateArray())
{
    string key = $"{node.GetProperty("type").GetString()}|{node.GetProperty("title").GetString()}|{node.GetProperty("parentId").GetString()}";
    actualSet.Add(key);
}

Console.WriteLine("Nodes in actual but not in expected:");
foreach (var key in actualSet)
{
    if (!expectedSet.Contains(key))
        Console.WriteLine($"  {key}");
}

Console.WriteLine("\nNodes in expected but not in actual:");
foreach (var key in expectedSet)
{
    if (!actualSet.Contains(key))
        Console.WriteLine($"  {key}");
}
