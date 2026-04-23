namespace TemplateParser.Core;
using DocumentFormat.OpenXml.Packaging;         // Needed for WordProcessingDocument.
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

// Provides functionality to parse a DOCX template and extract a hierarchical structure of headings as nodes.
public sealed class DocxParser
{

    // Parses a DOCX template file and returns a ParserResult containing a list of nodes representing the document's heading structure.

    /// <param name="filePath">The path to the DOCX file to parse.</param>
    /// <param name="templateId">The unique identifier for the template being parsed.</param>
    /// <returns>A ParserResult containing the extracted nodes.</returns>
    public ParserResult ParseDocxTemplate(string filePath, Guid templateId)
    {
        // Maps Word heading styles (e.g., Heading1, Heading2) to node types and their hierarchical levels.
        var headingMap = new Dictionary<string, (string nodeType, int level)>
        {
            { "Heading1", ("section", 1) },        // Top-level section
            { "Heading2", ("subsection", 2) },     // Second-level subsection
            { "Heading3", ("subsubsection", 3) }   // Third-level subsection
        };

        // List to store all parsed nodes.
        var nodes = new List<Node>();
        // DEBUG: Log file to trace node emission
        var debugLogPath = System.IO.Path.Combine(Environment.CurrentDirectory, "parse_debug.log");
        using var debugLog = new System.IO.StreamWriter(debugLogPath, append: false);
        void Log(string msg) { debugLog.WriteLine($"[{DateTime.Now:HH:mm:ss}] {msg}"); debugLog.Flush(); }

        // Stack to keep track of the current hierarchy while parsing headings.
        // Each entry contains: (node, heading level, sibling order)
        var stack = new Stack<(Node node, int level, int siblingIndex)>();

        // Tracks the next sibling order index for each parent node (by parentId).
        var siblingOrder = new Dictionary<Guid, int>();

        // --- Heuristic Engine Preparation ---
        // Pass 1: Collect all font sizes to determine baseline (body text) font size
        List<int> allFontSizes = new List<int>();
        using (WordprocessingDocument wordProcessingDocument = WordprocessingDocument.Open(filePath, false))
        {
            Body? body = wordProcessingDocument?.MainDocumentPart?.Document?.Body;
            ArgumentNullException.ThrowIfNull(body, "Document is empty.");
            foreach (OpenXmlElement element in body.Elements())
            {
                if (element is Paragraph para)
                {
                    foreach (var run in para.Elements<Run>())
                    {
                        var sz = run.RunProperties?.FontSize?.Val;
                        if (sz != null && int.TryParse(sz, out int size))
                        {
                            allFontSizes.Add(size);
                        }
                    }
                }
            }
        }
        // Compute baseline font size (mode)
        int baselineFontSize = allFontSizes.Count > 0 ? allFontSizes.GroupBy(x => x).OrderByDescending(g => g.Count()).First().Key : 22;
        // Initialize heuristic detector
        var heuristicDetector = new HeuristicHeadingDetector(baselineFontSize);

        // Open the DOCX file for reading (read-only mode).
        using (WordprocessingDocument wordProcessingDocument = WordprocessingDocument.Open(filePath, false))
        {
            // Get the document body. Throws if the document is empty.
            Body? body = wordProcessingDocument?.MainDocumentPart?.Document?.Body;
            ArgumentNullException.ThrowIfNull(body, "Document is empty.");

            // Use a for-loop to allow skipping elements (for lists)
            var bodyElements = body.Elements().ToList();
            for (int i = 0; i < bodyElements.Count; i++)
            {
                var element = bodyElements[i];
                // --- Paragraph Handling: Only one node per paragraph ---
                if (element is Paragraph p)
                {
                    string? style = p?.ParagraphProperties?.ParagraphStyleId?.Val;
                    string? text = p?.InnerText;
                    Log($"[PARA] i={i} style={style ?? "<none>"} text='{text?.Replace("\n", " ").Replace("\r", " ").Trim()}'.");
                    // Skip paragraphs that are empty, whitespace, or only contain non-visible characters
                    if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(text.Trim('\r','\n','\t',' '))) {
                        Log($"  Skipped: empty or whitespace.");
                        continue;
                    }
                    text = text.Trim();

                    // 1. List: If paragraph is a list item, skip here (handled in list logic below)
                    if (p.ParagraphProperties?.NumberingProperties != null) {
                        Log($"  Skipped: list item (handled in list logic).");
                        continue;
                    }

                    // 2. Image: If paragraph contains an image, skip here (handled in image logic below)
                    if (p.Descendants<Drawing>().Any()) {
                        Log($"  Skipped: contains image (handled in image logic).");
                        continue;
                    }

                    // 3. Heading: Style-based or heuristic heading detection
                    bool isHeading = false;
                    string nodeType = "paragraph";
                    int level = 0;
                    Dictionary<string, object> heuristics = new();
                    if (!string.IsNullOrWhiteSpace(style) && headingMap.TryGetValue(style, out var map))
                    {
                        nodeType = map.nodeType;
                        level = map.level;
                        isHeading = true;
                        heuristics["style"] = style;
                        Log($"  Detected heading by style: {nodeType} (level {level})");
                    }
                    else
                    {
                        string? inferred = heuristicDetector.InferHeadingLevel(p);
                        if (inferred != null)
                        {
                            isHeading = true;
                            nodeType = inferred;
                            level = nodeType == "section" ? 1 : nodeType == "subsection" ? 2 : 3;
                            heuristics["heuristic"] = true;
                            Log($"  Detected heading by heuristic: {nodeType} (level {level})");
                        }
                    }

                    if (isHeading)
                    {
                        // Maintain heading hierarchy using a stack
                        while (stack.Count > 0 && stack.Peek().level >= level)
                        {
                            stack.Pop();
                        }
                        Guid? parentId = stack.Count > 0 ? stack.Peek().node.Id : null;
                        int orderIndex = 0;
                        if (parentId.HasValue)
                        {
                            if (!siblingOrder.ContainsKey(parentId.Value)) siblingOrder[parentId.Value] = 0;
                            orderIndex = siblingOrder[parentId.Value]++;
                        }
                        else
                        {
                            if (!siblingOrder.ContainsKey(Guid.Empty)) siblingOrder[Guid.Empty] = 0;
                            orderIndex = siblingOrder[Guid.Empty]++;
                        }
                        var node = new Node
                        {
                            Id = Guid.NewGuid(),
                            TemplateId = templateId,
                            ParentId = parentId,
                            Type = nodeType,
                            Title = text,
                            OrderIndex = orderIndex,
                            MetadataJson = System.Text.Json.JsonSerializer.Serialize(heuristics)
                        };
                        nodes.Add(node);
                        stack.Push((node, level, orderIndex));
                        Log($"  Emitted heading node: {nodeType} '{text}' parent={parentId} order={orderIndex}");
                        continue; // Do not emit a text node for headings
                    }

                    // 4. Otherwise, emit a text node (not heading, not list, not image)
                    Guid? textParentId = stack.Count > 0 ? stack.Peek().node.Id : null;
                    int textOrderIndex = 0;
                    if (textParentId.HasValue)
                    {
                        if (!siblingOrder.ContainsKey(textParentId.Value)) siblingOrder[textParentId.Value] = 0;
                        textOrderIndex = siblingOrder[textParentId.Value]++;
                    }
                    else
                    {
                        if (!siblingOrder.ContainsKey(Guid.Empty)) siblingOrder[Guid.Empty] = 0;
                        textOrderIndex = siblingOrder[Guid.Empty]++;
                    }
                    var textNode = new Node
                    {
                        Id = Guid.NewGuid(),
                        TemplateId = templateId,
                        ParentId = textParentId,
                        Type = "text",
                        Title = "Text",
                        OrderIndex = textOrderIndex,
                        MetadataJson = System.Text.Json.JsonSerializer.Serialize(new { defaultText = text })
                    };
                    nodes.Add(textNode);
                    Log($"  Emitted text node: '{text}' parent={textParentId} order={textOrderIndex}");
                }
                else if (element is Table tbl)
                // --- Table Extraction ---
                // ...existing code...
                {
                    // ...existing code for table node emission...
                }
                // ...existing code for table, image, and list node emission as before revert...
            }
        }

        // Sort nodes by parentId and orderIndex for deterministic output order.
        nodes = nodes.OrderBy(n => n.ParentId ?? Guid.Empty).ThenBy(n => n.OrderIndex).ToList();

        // Return the result containing all parsed nodes.
        return new ParserResult { Nodes = nodes };
    }
}
