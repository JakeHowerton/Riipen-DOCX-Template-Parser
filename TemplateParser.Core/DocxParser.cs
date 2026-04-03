using System;
using DocumentFormat.OpenXml.Packaging;         // Needed for WordProcessingDocument.
using DocumentFormat.OpenXml.Wordprocessing;    // Needed for all Word schema objects (Body, Paragraph, etc.)


namespace TemplateParser.Core;

public sealed class DocxParser
{
    public ParserResult ParseDocxTemplate(string filePath, Guid templateId)
    {
        var nodes = new List<Node>();
        // Tracking the last node added at each level (0 = Root, 1 = H1, 2 = H2, etc.)
        var lastNodesAtLevel = new Dictionary<int, Node>();
        int globalOrderIndex = 0;

        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
        {
            Body? body = wordDoc?.MainDocumentPart?.Document?.Body;
            ArgumentNullException.ThrowIfNull(body, "The document body could not be loaded.");

            foreach (Paragraph p in body.Descendants<Paragraph>())
            {
                string text = p.InnerText.Trim();
                if (string.IsNullOrWhiteSpace(text)) continue;

                string style = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "Normal";

                // Deconstruct the tuple from our new helper
                var (level, type) = GetNodeDetails(style);

                var newNode = new Node
                {
                    Id = Guid.NewGuid(), // Globally Unique Identifier
                    TemplateId = templateId, // Using the passed-in argument
                    Title = text,
                    Type = type, // Now sets "Section", "Subsection", etc.
                    OrderIndex = globalOrderIndex++,
                    MetadataJson = "{}",
                };
                // 3. Assign Parent Logic
                if (level == 1)
                {
                    newNode.ParentId = null; // Root level
                }
                else
                {
                    // Find the closest ancestor (Exp: H3 looks for H2, then H1)
                    newNode.ParentId = FindParentId(level, lastNodesAtLevel);
                }

                // 4. Update state and results
                lastNodesAtLevel[level] = newNode;
                nodes.Add(newNode);
            }
        }

        return new ParserResult { Nodes = nodes };
    }

    private (int Level, string Type) GetNodeDetails(string styleId)
    {
        return styleId switch
        {
            "Heading1" => (1, "Section"),
            "Heading2" => (2, "Subsection"),
            "Heading3" => (3, "Subsubsection"),
            _ => (4, "Content") // Default for Normal text or other styles
        };
    }
    private Guid? FindParentId(int currentLevel, Dictionary<int, Node> lastNodes)
    {
        // Search backwards from the current level to find the nearest parent
        for (int i = currentLevel - 1; i >= 1; i--)
        {
            if (lastNodes.ContainsKey(i)) return lastNodes[i].Id;
        }
        return null;
    }
}





//To-Do Section
// TODO (Week 1-4): Implement core DOCX parsing here.
// Recommended responsibilities for this method:
// 1) [Week 1] Learn DOCX structure and print paragraphs from the document.
// (Completed Milestone 1 with what we did with sherbert in class)
// 2) [Week 2] Build section hierarchy using Word heading styles.
// (This is what we are working on now within milestone 2)


// 3) [Week 3] Detect tables, lists, and images as structured content nodes.
// 4) [Week 4] Add formatting heuristics for files missing heading styles.
// 5) [Week 2-4] Create Node instances with:
//    - Id: new Guide for each node
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

//throw new NotImplementedException("DOCX parsing is intentionally not implemented in this starter repository.");
