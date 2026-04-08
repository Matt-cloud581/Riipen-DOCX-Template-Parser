namespace TemplateParser.Core;
using DocumentFormat.OpenXml.Packaging;         // Needed for WordProcessingDocument.
using DocumentFormat.OpenXml.Wordprocessing;

public sealed class DocxParser
{
    public ParserResult ParseDocxTemplate(string filePath, Guid templateId)
    {
        // Map Word heading styles to node types
        var headingMap = new Dictionary<string, (string nodeType, int level)>
        {
            { "Heading1", ("section", 1) },
            { "Heading2", ("subsection", 2) },
            { "Heading3", ("subsubsection", 3) }
        };

        var nodes = new List<Node>();
        // Stack: (node, heading level, sibling order)
        var stack = new Stack<(Node node, int level, int siblingIndex)>();
        var siblingOrder = new Dictionary<Guid, int>(); // parentId -> next orderIndex

        using (WordprocessingDocument wordProcessingDocument = WordprocessingDocument.Open(filePath, false))
        {
            Body? body = wordProcessingDocument?.MainDocumentPart?.Document?.Body;
            ArgumentNullException.ThrowIfNull(body, "Document is empty.");

            foreach (Paragraph p in body.Descendants<Paragraph>())
            {
                string? style = p?.ParagraphProperties?.ParagraphStyleId?.Val;
                string? text = p?.InnerText?.Trim();
                if (string.IsNullOrWhiteSpace(style) || string.IsNullOrWhiteSpace(text))
                    continue;

                // Normalize style name (Word can use Heading1, Heading2, etc.)
                if (!headingMap.TryGetValue(style, out var map))
                    continue;

                var nodeType = map.nodeType;
                var level = map.level;

                // Pop stack until we find the parent for this heading level
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
                    // Root node
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
                    MetadataJson = "{}"
                };
                nodes.Add(node);
                stack.Push((node, level, orderIndex));
            }
        }

        // Sort nodes by parentId, then orderIndex for deterministic output
        nodes = nodes.OrderBy(n => n.ParentId ?? Guid.Empty).ThenBy(n => n.OrderIndex).ToList();
        return new ParserResult { Nodes = nodes };
    }
}
