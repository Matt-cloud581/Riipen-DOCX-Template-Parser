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

            // Iterate through all direct children of the document body (preserves order of tables, paragraphs, etc.)
            foreach (OpenXmlElement element in body.Elements())
            {
                // --- Heading Extraction ---
                // Detect paragraphs styled as headings and emit section/subsection nodes
                if (element is Paragraph p)
                {
                    string? style = p?.ParagraphProperties?.ParagraphStyleId?.Val;
                    string? text = p?.InnerText?.Trim();
                    if (string.IsNullOrWhiteSpace(text))
                        continue;
                    bool isHeading = false;
                    string nodeType = "paragraph";
                    int level = 0;
                    Dictionary<string, object> heuristics = new();
                    // --- Style-based detection ---
                    if (!string.IsNullOrWhiteSpace(style) && headingMap.TryGetValue(style, out var map))
                    {
                        nodeType = map.nodeType;
                        level = map.level;
                        isHeading = true;
                        heuristics["style"] = style;
                    }
                    else
                    {
                        // --- Heuristic detection (unified pipeline) ---
                        string? inferred = heuristicDetector.InferHeadingLevel(p);
                        if (inferred != null)
                        {
                            isHeading = true;
                            nodeType = inferred;
                            // Infer level from nodeType
                            level = nodeType == "section" ? 1 : nodeType == "subsection" ? 2 : 3;
                            heuristics["heuristic"] = true;
                        }
                    }
                    if (isHeading)
                    {
                        // Maintain heading hierarchy using a stack
                        while (stack.Count > 0 && stack.Peek().level >= level)
                        {
                            stack.Pop();
                        }
                        // Determine the parent node's ID (if any)
                        Guid? parentId = stack.Count > 0 ? stack.Peek().node.Id : null;
                        int orderIndex = 0;
                        // Assign the correct sibling order index for this node under its parent
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
                        // Create a new node for this heading
                        var node = new Node
                        {
                            Id = Guid.NewGuid(),           // Unique identifier for the node
                            TemplateId = templateId,        // The template this node belongs to
                            ParentId = parentId,            // The parent node's ID (null for root)
                            Type = nodeType,                // Node type (section, subsection, etc.)
                            Title = text,                   // The heading text
                            OrderIndex = orderIndex,        // Sibling order under the parent
                            MetadataJson = System.Text.Json.JsonSerializer.Serialize(heuristics)
                        };
                        // Add the node to the list and push it onto the stack for hierarchy tracking
                        nodes.Add(node);
                        stack.Push((node, level, orderIndex));
                        continue;
                    }
                }
                else if (element is Table tbl)
                // --- Table Extraction ---
                // Detect tables, extract rows/columns/cell text, and emit table node with metadata
                {
                    // Get all rows in the table
                    var rows = tbl.Elements<TableRow>().ToList();
                    int rowCount = rows.Count;
                    // Assume all rows have the same number of columns as the first row
                    int colCount = rows.FirstOrDefault()?.Elements<TableCell>().Count() ?? 0;
                    // Build a 2D array of cell text
                    var tableData = new List<List<string>>();
                    foreach (var row in rows)
                    {
                        var rowData = new List<string>();
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            // Concatenate all text in the cell
                            string cellText = string.Join(" ", cell.Descendants<Text>().Select(t => t.Text).Where(t => !string.IsNullOrWhiteSpace(t)));
                            rowData.Add(cellText);
                        }
                        tableData.Add(rowData);
                    }
                    // Build metadata JSON for the table node
                    var metadata = new {
                        rows = rowCount,
                        columns = colCount,
                        tableData = tableData
                    };
                    string metadataJson = System.Text.Json.JsonSerializer.Serialize(metadata);
                    // Determine parent and order index for the table node
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
                    // Create and add the table node
                    var tableNode = new Node
                    {
                        Id = Guid.NewGuid(),
                        TemplateId = templateId,
                        ParentId = parentId,
                        Type = "table",
                        Title = string.Empty,
                        OrderIndex = orderIndex,
                        MetadataJson = metadataJson
                    };
                    nodes.Add(tableNode);
                }
                else if (element is Paragraph paraWithDrawing && paraWithDrawing.Descendants<Drawing>().Any())
                // --- Image Extraction ---
                // Detect paragraphs containing Drawing elements (images), extract EMU dimensions, and emit image node
                {
                    var drawing = paraWithDrawing.Descendants<Drawing>().FirstOrDefault();
                    if (drawing != null)
                    {
                        // Try to get EMU dimensions from wp:extent
                        var extent = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>().FirstOrDefault()?.Extent;
                        long widthEmu = 0;
                        long heightEmu = 0;
                        if (extent != null)
                        {
                            widthEmu = extent.Cx != null ? extent.Cx.Value : 0;
                            heightEmu = extent.Cy != null ? extent.Cy.Value : 0;
                        }
                        // Optional: title/description extraction (not always present)
                        string? title = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().FirstOrDefault()?.Title;
                        string? descr = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().FirstOrDefault()?.Description;
                        // Build metadata JSON for the image node
                        var imageMetadata = new {
                            widthEmu = widthEmu,
                            heightEmu = heightEmu,
                            title = title,
                            description = descr
                        };
                        string imageMetadataJson = System.Text.Json.JsonSerializer.Serialize(imageMetadata);
                        // Determine parent and order index for the image node
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
                        // Create and add the image node
                        var imageNode = new Node
                        {
                            Id = Guid.NewGuid(),
                            TemplateId = templateId,
                            ParentId = parentId,
                            Type = "image",
                            Title = title ?? string.Empty,
                            OrderIndex = orderIndex,
                            MetadataJson = imageMetadataJson
                        };
                        nodes.Add(imageNode);
                    }
                }
                else if (element is Paragraph paraList && paraList.ParagraphProperties?.NumberingProperties != null)
                // --- List Detection ---
                // Detect paragraphs with numbering properties, group consecutive list items, and emit a list node
                {
                    var listItems = new List<string>();
                    var listType = "unknown";
                    var currentElement = element;
                    var bodyElements = body.Elements().ToList();
                    int startIdx = bodyElements.IndexOf(element);
                    int idx = startIdx;
                    // Group consecutive paragraphs that are part of the same list
                    while (idx < bodyElements.Count)
                    {
                        var el = bodyElements[idx];
                        if (el is Paragraph para && para.ParagraphProperties?.NumberingProperties != null)
                        {
                            string itemText = para.InnerText.Trim();
                            if (!string.IsNullOrWhiteSpace(itemText))
                                listItems.Add(itemText);
                            // Determine list type (bullet or decimal)
                            var numProp = para.ParagraphProperties.NumberingProperties;
                            var numId = numProp.NumberingId?.Val;
                            var lvl = numProp.NumberingLevelReference?.Val;
                            // Heuristic: if numId is present, try to infer type
                            if (listType == "unknown")
                            {
                                // If the paragraph has a numbering format, try to get it
                                var numberingPart = wordProcessingDocument.MainDocumentPart?.NumberingDefinitionsPart;
                                if (numberingPart != null && numId != null)
                                {
                                    var num = numberingPart.Numbering?.Elements<NumberingInstance>().FirstOrDefault(n => n.NumberID == numId);
                                    var absNumId = num?.AbstractNumId?.Val;
                                    var absNum = numberingPart.Numbering?.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId == absNumId);
                                    var lvlEl = absNum?.Elements<Level>().FirstOrDefault(l => l.LevelIndex == lvl);
                                    var format = lvlEl?.NumberingFormat?.Val;
                                    if (format != null)
                                    {
                                        listType = format.Value == NumberFormatValues.Bullet ? "bullet" : "numbered";
                                    }
                                }
                            }
                                    // --- Text Classification ---
                                    // For paragraphs not handled above, classify as sentence or paragraph node
                                    else if (element is Paragraph paraText)
                                    {
                                        // Skip empty paragraphs
                                        string text = paraText.InnerText.Trim();
                                        if (string.IsNullOrWhiteSpace(text))
                                            continue;
                                        // Skip if this is a heading (already handled)
                                        string? style = paraText.ParagraphProperties?.ParagraphStyleId?.Val;
                                        if (!string.IsNullOrWhiteSpace(style) && headingMap.ContainsKey(style))
                                            continue;
                                        // Skip if this is a list (already handled)
                                        if (paraText.ParagraphProperties?.NumberingProperties != null)
                                            continue;
                                        // Skip if this contains an image (already handled)
                                        if (paraText.Descendants<Drawing>().Any())
                                            continue;

                                        // --- Heuristics for text classification ---
                                        // Sentence: short text, no terminal block cues, single run
                                        // Paragraph: longer, multiple sentences, trailing punctuation patterns
                                        bool isSentence = false;
                                        // Heuristic: consider a sentence if text is short and has only one sentence-ending punctuation
                                        int wordCount = text.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length;
                                        int periodCount = text.Count(c => c == '.' || c == '!' || c == '?');
                                        // Consider a sentence if <= 20 words and <= 1 period/exclamation/question mark
                                        if (wordCount <= 20 && periodCount <= 1)
                                            isSentence = true;
                                        // Otherwise, treat as paragraph
                                        string nodeType = isSentence ? "sentence" : "paragraph";
                                        // Build metadata JSON (can be extended with more cues)
                                        var textMetadata = new {
                                            length = wordCount,
                                            punctuationCount = periodCount
                                        };
                                        string textMetadataJson = System.Text.Json.JsonSerializer.Serialize(textMetadata);
                                        // Determine parent and order index for the text node
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
                                        // Create and add the text node
                                        var textNode = new Node
                                        {
                                            Id = Guid.NewGuid(),
                                            TemplateId = templateId,
                                            ParentId = textParentId,
                                            Type = nodeType,
                                            Title = text,
                                            OrderIndex = textOrderIndex,
                                            MetadataJson = textMetadataJson
                                        };
                                        nodes.Add(textNode);
                                    }
                            idx++;
                        }
                        else
                        {
                            break;
                        }
                    }
                    // Build metadata JSON for the list node
                    var listMetadata = new {
                        listType = listType,
                        items = listItems
                    };
                    string listMetadataJson = System.Text.Json.JsonSerializer.Serialize(listMetadata);
                    // Determine parent and order index for the list node
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
                    // Create and add the list node
                    var listNode = new Node
                    {
                        Id = Guid.NewGuid(),
                        TemplateId = templateId,
                        ParentId = parentId,
                        Type = "list",
                        Title = string.Empty,
                        OrderIndex = orderIndex,
                        MetadataJson = listMetadataJson
                    };
                    nodes.Add(listNode);
                }
            }
        }

        // Sort nodes by parentId and orderIndex for deterministic output order.
        nodes = nodes.OrderBy(n => n.ParentId ?? Guid.Empty).ThenBy(n => n.OrderIndex).ToList();

        // Return the result containing all parsed nodes.
        return new ParserResult { Nodes = nodes };
    }
}
