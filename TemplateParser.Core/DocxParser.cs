namespace TemplateParser.Core;
using DocumentFormat.OpenXml.Packaging;         // Needed for WordProcessingDocument.
using DocumentFormat.OpenXml.Wordprocessing;

public sealed class DocxParser
{
    public ParserResult ParseDocxTemplate(string filePath, Guid templateId)
    {
        {
    // 1. Open the word document in read mode.
    // 2. Parse the document.xml into XML object using the DocumentFormat.OpenXml library.
    using (WordprocessingDocument wordProcessingDocument = WordprocessingDocument.Open(filePath, false))
    {
        // The original line we wrote in class:
        // Body body = wordProcessingDocument.MainDocumentPart.Document.Body;
        
        // A more robust version that fails gracefully if the document is not structured properly:
        Body? body = wordProcessingDocument?.MainDocumentPart?.Document?.Body;
        ArgumentNullException.ThrowIfNull(body, "Document is empty.");
        
        // 3. Loop through every paragraph.
        foreach (Paragraph p in body.Descendants<Paragraph>())
        {
            // 4. Extract and display the paragraph style.
            // The original line we wrote in class:
            // string style = p?.ParagraphProperties?.ParagraphStyleId?.Val;
            // A more robust version:
            string? style = p?.ParagraphProperties?.ParagraphStyleId?.Val ?? "No Style";
            Console.WriteLine(style);

            // 5. Extract and display the actual text.
            string? text = p?.InnerText;
            Console.WriteLine(text);

            // I added the following line to space the output out a little better:
            Console.WriteLine("--------------------------------");
        }
    }
        }
        // TODO (Week 1-4): Implement core DOCX parsing here.
        // Recommended responsibilities for this method:
        // 1) [Week 1] Learn DOCX structure and print paragraphs from the document.
        // 2) [Week 2] Build section hierarchy using Word heading styles.
        // 3) [Week 3] Detect tables, lists, and images as structured content nodes.
        // 4) [Week 4] Add formatting heuristics for files missing heading styles.
        // 5) [Week 2-4] Create Node instances with:
        //    - Id: new Guid for each node
        //    - TemplateId: the templateId argument
        //    - ParentId: null for root nodes, set for child nodes
        //    - Type/Title/OrderIndex/MetadataJson based on parsed content
        // 6) [Week 4] Return ParserResult with Nodes in deterministic order.
        //
        // Helper guidance [Week 3-6]:
        // - YES, create helper classes if this method gets long or hard to read.
        // - Keep helpers inside TemplateParser.Core (for example, Parsing/ or Utilities/ folders).
        // - Keep this method as the high-level orchestration entry point.
        // - In Week 6, refactor large blocks from this method into focused helper classes.
        //
        // Do not place parsing logic in the CLI project; keep it in Core.
        throw new NotImplementedException("DOCX parsing is intentionally not implemented in this starter repository.");
    }
}
