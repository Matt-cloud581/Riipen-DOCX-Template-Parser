
# TemplateParser.NET

TemplateParser.NET is a .NET 10 solution for extracting structured content from DOCX templates. It is designed for educational use, with a focus on clarity, extensibility, and robust handling of real-world documents.

## Parsing Strategy

The parser reads DOCX files using the OpenXML SDK (no Office/Word required). It processes the document body sequentially, mapping each content block (paragraph, table, list, image) to a `Node` object. Hierarchy is inferred using heading styles and heuristics, with parent-child relationships built via a stack.

- **Headings:** Detected by Word styles (Heading1/2/3) or heuristics (font size, boldness).
- **Paragraphs:** Non-heading, non-list, non-image paragraphs become `text` nodes.
- **Lists:** Numbered/bulleted paragraphs are grouped as `list` nodes.
- **Tables:** Each table is a `table` node with metadata.
- **Images:** Inline images are extracted as `image` nodes.

## Heading Detection Heuristics

- **Word Styles:** Heading1 → `section`, Heading2 → `subsection`, Heading3 → `subsubsection`.
- **Font Size:** Paragraphs with font size above the document mode are candidates.
- **Boldness:** Bold, large text is more likely a heading.
- **Fallback:** If no style, use heuristics to infer heading level.

**Example:**
- "Introduction" in Heading1 style → `section` node
- Large, bold paragraph with no style → `section` (heuristic)

## How to Run the CLI

From the repo root:

```sh
dotnet run --project TemplateParser.Cli -- parse sample-documents/sample.docx 00000000-0000-0000-0000-000000000000
```

- Output: `output.json` in the working directory
- Arguments: `<filePath> <templateId>`
- Example output: JSON array of nodes

## Integration Instructions

To use the parser in your own .NET project:

1. Reference `TemplateParser.Core`.
2. Call:

```csharp
var parser = new DocxParser();
var result = parser.ParseDocxTemplate("path/to/file.docx", Guid.NewGuid());
foreach (var node in result.Nodes) { /* ... */ }
```
3. Ensure `DocumentFormat.OpenXml` NuGet package is installed.

## Known Limitations

- Does not handle nested tables or floating images.
- Custom numbering/bullet styles may not be detected as lists.
- Heuristic heading detection may misclassify unusual formatting.
- Only DOCX files are supported (not DOC, PDF, etc.).
- No support for tracked changes or comments.

## Dependency Hygiene

- No Microsoft Office, Word Interop, or Office Automation required.
- All dependencies are managed via NuGet (`DocumentFormat.OpenXml`).
- Runs on any machine with .NET 10+.

## Solution Structure Checklist

- `TemplateParser.sln` (solution file)
- `TemplateParser.Core` (core logic)
- `TemplateParser.Cli` (command-line interface)
- `TemplateParser.Tests` (xUnit tests)
- `sample-documents/` (input DOCX files)
- `expected/` (expected output JSON)
- All packages restored via `dotnet restore`, no manual DLLs
- No Office/Word required

---

For more details, see code comments and test cases in the repository.

## Suggested 6-Week Path

- **Week 1:** Understand the `.docx` file structure and print paragraphs from the document
- **Week 2:** Build the basic section hierarchy from Word heading styles
- **Week 3:** Detect tables, lists, and images as structured content nodes
- **Week 4:** Implement formatting heuristics when heading styles are missing
- **Week 5:** Wire the parser into a CLI tool and add integration tests
- **Week 6:** Documentation, refactoring, and final delivery

In every week, ask: "How does this step improve the quality of the `Node` objects I produce?"

## Run the CLI

From the repo root:

```bash
dotnet run --project TemplateParser.Cli -- parse <filePath> <templateId>
```

Shorter option from the repo root (recommended):

```bash
./parse <filePath> <templateId>
```

From inside `TemplateParser.Cli`:

```bash
dotnet run -- parse <filePath> <templateId>
```

## Learning Workflow (Recommended)

1. Add a sample DOCX in `sample-documents`
2. Write or update tests in `TemplateParser.Tests`
3. Implement parser behavior in `TemplateParser.Core`
4. Run:
   - `dotnet build`
   - `dotnet test`
5. Use CLI to inspect JSON output manually
6. Save expected outputs in `expected` and compare

## Important Note

Parser implementation is intentionally incomplete in this starter repository. You are expected to implement:

- DOCX reading/parsing
- node extraction
- relationship building
- metadata handling
- test coverage
