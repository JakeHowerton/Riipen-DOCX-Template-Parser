using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging; // Needed for WordProcessingDocument.
using DocumentFormat.OpenXml.Wordprocessing; // Needed for all Word schema objects (Body, Paragraph, etc.)

// Alias for Drawing to avoid conflicts with Wordprocessing.Drawing
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace TemplateParser.Core;

public class HeadingInference
{
    public int? Level { get; set; }
    public Dictionary<string, object> Formatting { get; set; } = new();
}

public class HeuristicHeadingDetector
{
    private const double DefaultFontSize = 11.0;

    public HeadingInference InferHeadingLevel(Paragraph p)
    {
        var inference = new HeadingInference();
        string text = p.InnerText.Trim();

        // Basic filter: if it's empty or way too long, it's just content
        if (string.IsNullOrWhiteSpace(text) || text.Length > 150) return inference;

        var runs = p.Descendants<Run>().ToList();

        // Signal 1: Font Size
        double fontSize = GetMaxFontSize(runs);

        // Signal 2: Bold & Italic Weight
        bool isBold = runs.Any(r =>
            r.RunProperties?.Bold != null ||
            r.RunProperties?.RunStyle?.Val?.Value?.Contains("Bold", StringComparison.OrdinalIgnoreCase) == true);

        bool isItalic = runs.Any(r => r.RunProperties?.Italic != null);

        // Signal 3: Numbering Patterns
        bool hasNumbering = Regex.IsMatch(text, @"^(\d+(\.\d+)*|[A-Z]\.|Section\s\d+)", RegexOptions.IgnoreCase);

        // Signal 4: Spacing
        var spacing = p.ParagraphProperties?.SpacingBetweenLines;
        bool hasSpaceBefore = spacing?.Before != null;

        // Populate Formatting Metadata for the Output
        inference.Formatting["fontSize"] = fontSize;
        inference.Formatting["isBold"] = isBold;
        inference.Formatting["isItalic"] = isItalic;
        inference.Formatting["hasNumbering"] = hasNumbering;

        // Weighted Scoring Logic
        double score = 0;
        if (fontSize > DefaultFontSize + 4) score += 3;      // Large font
        else if (fontSize > DefaultFontSize + 1) score += 1.5; // Med font

        if (isBold) score += 2;
        if (isItalic) score += 0.5;
        if (hasNumbering) score += 2.5;
        if (hasSpaceBefore) score += 1;

        // Inference Thresholds
        if (score >= 6) inference.Level = 1;      // Section
        else if (score >= 4) inference.Level = 2; // Subsection
        else if (score >= 2.5) inference.Level = 3; // Subsubsection

        return inference;
    }

    private double GetMaxFontSize(IEnumerable<Run> runs)
    {
        var sizes = runs.Select(r =>
        {
            var sz = r.RunProperties?.FontSize?.Val?.Value;
            if (sz != null && double.TryParse(sz, out double val)) return val / 2.0;
            return DefaultFontSize;
        });
        return sizes.Any() ? sizes.Max() : DefaultFontSize;
    }
}

public sealed class DocxParser
{
    private readonly HeuristicHeadingDetector _heuristics = new();

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

            // Flattening the structure to process Paragraphs and Tables in sequence
            var elements = body.ChildElements.Where(e => e is Paragraph || e is Table);
            List<Paragraph> listBuffer = new();

            foreach (var element in elements)
            {
                // Handle List Grouping: Flush buffer if current element is not a list item
                if (element is not Paragraph pList || pList.ParagraphProperties?.NumberingProperties == null)
                {
                    FlushListBuffer(listBuffer, nodes, lastNodesAtLevel, ref globalOrderIndex, templateId);
                }

                if (element is Paragraph p)
                {
                    // Check for Images first (often nested in Paragraphs)
                    var drawing = p.Descendants<Drawing>().FirstOrDefault();
                    if (drawing != null)
                    {
                        nodes.Add(CreateImageNode(drawing, templateId, ref globalOrderIndex, lastNodesAtLevel));
                        continue;
                    }

                    // Check for Lists
                    if (p.ParagraphProperties?.NumberingProperties != null)
                    {
                        listBuffer.Add(p);
                        continue;
                    }

                    // Standard Text / Heading Processing
                    if (string.IsNullOrWhiteSpace(p.InnerText)) continue;
                    nodes.Add(ProcessParagraph(p, templateId, ref globalOrderIndex, lastNodesAtLevel));
                }
                else if (element is Table table)
                {
                    nodes.Add(CreateTableNode(table, templateId, ref globalOrderIndex, lastNodesAtLevel));
                }
            }
            // Final flush for trailing lists
            FlushListBuffer(listBuffer, nodes, lastNodesAtLevel, ref globalOrderIndex, templateId);
        }

        return new ParserResult { Nodes = nodes };
    }

    private Node ProcessParagraph(Paragraph p, Guid templateId, ref int orderIndex, Dictionary<int, Node> lastNodes)
    {
        string text = p.InnerText.Trim();
        string styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "";

        // 1. Try Style Detection
        // Deconstruct the tuple from our new helper
        var (level, type) = GetNodeDetailsFromStyle(styleId);

        // 2. Run Heuristics
        var inference = _heuristics.InferHeadingLevel(p);

        // Fallback to Heuristics if it's tagged as Content but looks like a Header
        if (level == 4 && inference.Level.HasValue)
        {
            level = inference.Level.Value;
            type = level switch { 1 => "Section", 2 => "Subsection", _ => "Subsubsection" };
        }

        // Classify Text: Sentence vs Paragraph
        if (type == "Content")
        {
            type = (text.Contains('.') && text.Length > 150) ? "Paragraph" : "Sentence";
        }

        var node = new Node
        {
            Id = Guid.NewGuid(),
            TemplateId = templateId,
            Title = text.Length > 50 ? text.Substring(0, 47) + "..." : text,
            Type = type,
            OrderIndex = orderIndex++,
            ParentId = level == 1 ? null : FindParentId(level, lastNodes),
            // Formatting signals now show up in MetadataJson
            MetadataJson = JsonSerializer.Serialize(inference.Formatting) //Formating of bold, underlined, etc.
        };

        // Update tracking and CLEAR lower-level children to maintain hierarchy integrity
        lastNodes[level] = node;
        for (int i = level + 1; i <= 4; i++)
        {
            if (lastNodes.ContainsKey(i)) lastNodes.Remove(i);
        }

        return node;
    }

    private (int Level, string Type) GetNodeDetailsFromStyle(string styleId) => styleId switch
    {
        "Heading1" => (1, "Section"),
        "Heading2" => (2, "Subsection"),
        "Heading3" => (3, "Subsubsection"),
        _ => (4, "Content") // Default for Normal text or other styles
    };

    private Node CreateTableNode(Table table, Guid templateId, ref int orderIndex, Dictionary<int, Node> lastNodes)
    {
        var rows = table.Elements<TableRow>();
        var tableData = rows.Select(row =>
            row.Elements<TableCell>().Select(c => c.InnerText.Trim()).ToList()
        ).ToList();

        return new Node
        {
            Id = Guid.NewGuid(), // Globally Unique Identifier
            TemplateId = templateId, // Using the passed-in argument
            Title = "Table Structure",
            Type = "Table",
            OrderIndex = orderIndex++,
            ParentId = FindParentId(4, lastNodes),
            MetadataJson = JsonSerializer.Serialize(new
            {
                rowCount = tableData.Count,
                colCount = tableData.FirstOrDefault()?.Count ?? 0,
                tableData = tableData
            })
            //If the content is anything that is not directly a header,subheader, or content. 
            // This gets filled in the with properties of the table
        };
    }

    private Node CreateImageNode(Drawing drawing, Guid templateId, ref int orderIndex, Dictionary<int, Node> lastNodes)
    {
        var extent = drawing.Descendants<Wp.Extent>().FirstOrDefault();
        return new Node
        {
            Id = Guid.NewGuid(), // Globally Unique Identifier
            TemplateId = templateId,
            Title = "Embedded Image",
            Type = "Image",
            OrderIndex = orderIndex++,
            ParentId = FindParentId(4, lastNodes),
            MetadataJson = JsonSerializer.Serialize(new
            {
                widthEmu = extent?.Cx?.Value ?? 0,
                heightEmu = extent?.Cy?.Value ?? 0
            })
            //if the content is anything that is not directly a header,subheader, or content. 
            // This gets filled in the with properties of the image
        };
    }

    private void FlushListBuffer(List<Paragraph> buffer, List<Node> nodes, Dictionary<int, Node> lastNodes, ref int orderIndex, Guid templateId)
    {
        if (buffer.Count == 0) return;

        var items = buffer.Select(p => p.InnerText.Trim()).ToList();
        var isNumbered = buffer.First().ParagraphProperties?.NumberingProperties?.NumberingId?.Val != null;

        nodes.Add(new Node
        {
            Id = Guid.NewGuid(), // Globally Unique Identifier
            TemplateId = templateId,
            Title = "List Content",
            Type = "List",
            OrderIndex = orderIndex++,
            ParentId = FindParentId(4, lastNodes),
            MetadataJson = JsonSerializer.Serialize(new
            {
                listType = isNumbered ? "Numbered" : "Bullet",
                items = items
            })
        });

        buffer.Clear();
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
