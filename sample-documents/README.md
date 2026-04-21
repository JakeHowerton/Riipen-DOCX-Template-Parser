Required README Sections

• Parsing Strategy: overall approach to structure inference

The parser employs a linear-to-hierarchical transformation strategy. It reads the document body sequentially and classifies elements into specific Node types while tracking depth to maintain parent-child relationships.
Sequence Processing: The parser flattens the document into a list of Paragraphs and Tables, processing them in their visual order to assign a deterministic OrderIndex.
Hierarchy Reconstruction: A tracking mechanism maintains the "last seen" node at each heading level (Section, Subsection, etc.). New nodes are automatically assigned as children of the nearest active parent level.
Node Classification: Elements are categorized into types: Section, Subsection, Paragraph, Sentence, Table, List, or Image

• Heading Detection Heuristics: each signal explained with examples

Font size: \u0022fontSize\u0022:11

Bold: \u0022isBold\u0022:true (if bolded True of False)

Italic: \u0022isItalic\u0022:true (if italic True of False)

Numbering: \u0022hasNumbering\u0022:true (if numbered True of False)

For images only:
Width: \u0022widthEmu\u0022:1928027 (width of image)
Height: \u0022heightEmu\u0022:739204 (height of image)

For lists only:
Types of list (bullet, numbered, etc): \u0022listType\u0022:\u0022Numbered\u0022
Contects: \u0022items\u0022:[\u0022Football\u0022,\u0022Basketball\u0022,\u0022People\u0022]

For Tables only:
Amount of rows: \u0022rowCount\u0022:3
Amount of columns: \u0022colCount\u0022:3
Table data: \u0022tableData\u0022:[[\u0022Insert 1\u0022,\u0022Insert 2\u0022,\u0022Insert 3\u0022],[\u0022Dolphins\u0022,\u0022Celtics\u0022,\u0022Tiger Woods\u0022],[\u0022Lions\u0022,\u0022Heats\u0022,\u0022Mike Tyson\u0022]]

• How to Run the CLI: step-by-step with example output

dotnet run --project TemplateParser.Cli -- parse "sample-documents/Sample Document.docx" "00000000-0000-0000-0000-000000000000"

{
"nodes": [
{
"id": "c2b9f7c8-f385-44c8-b1f9-741e183f8676",
"templateId": "00000000-0000-0000-0000-000000000000",
"type": "Section",
"title": "Section 1",
"orderIndex": 0,
"metadataJson": "{\u0022fontSize\u0022:11,\u0022isBold\u0022:false,\u0022isItalic\u0022:false,\u0022hasNumbering\u0022:true}"
},
{
"id": "c4f055e0-0fb6-48f9-8b05-49bc5e33d6e4",
"templateId": "00000000-0000-0000-0000-000000000000",
"parentId": "c2b9f7c8-f385-44c8-b1f9-741e183f8676",
"type": "Sentence",
"title": "This file is meant to be just a sample for writ...",
"orderIndex": 1,
"metadataJson": "{\u0022fontSize\u0022:11,\u0022isBold\u0022:false,\u0022isItalic\u0022:false,\u0022hasNumbering\u0022:false}"
},
{
"id": "7d86416c-535a-4400-84d6-2562a228832c",
"templateId": "00000000-0000-0000-0000-000000000000",
"parentId": "c2b9f7c8-f385-44c8-b1f9-741e183f8676",
"type": "Sentence",
"title": "The contents really don\u0027t matter.",
"orderIndex": 2,
"metadataJson": "{\u0022fontSize\u0022:11,\u0022isBold\u0022:true,\u0022isItalic\u0022:false,\u0022hasNumbering\u0022:false}"
},
{
"id": "7161b823-e7c2-4639-9580-f842c5a0bb76",
"templateId": "00000000-0000-0000-0000-000000000000",
"parentId": "c2b9f7c8-f385-44c8-b1f9-741e183f8676",
"type": "Sentence",
"title": "But we want to understand what is a \u0022paragraph\u0022...",
"orderIndex": 3,
"metadataJson": "{\u0022fontSize\u0022:11,\u0022isBold\u0022:false,\u0022isItalic\u0022:false,\u0022hasNumbering\u0022:false}"
},
{
"id": "6ba38402-f030-4721-a578-d605af2898bb",
"templateId": "00000000-0000-0000-0000-000000000000",
"parentId": "c2b9f7c8-f385-44c8-b1f9-741e183f8676",
"type": "Subsection",
"title": "Subsection",
"orderIndex": 4,
"metadataJson": "{\u0022fontSize\u0022:11,\u0022isBold\u0022:false,\u0022isItalic\u0022:false,\u0022hasNumbering\u0022:false}"
},
{
"id": "0822decb-7677-4db7-9590-0f9fd0dd4bf5",
"templateId": "00000000-0000-0000-0000-000000000000",
"parentId": "6ba38402-f030-4721-a578-d605af2898bb",
"type": "Sentence",
"title": "We also want to understand what is a style.",
"orderIndex": 5,
"metadataJson": "{\u0022fontSize\u0022:11,\u0022isBold\u0022:false,\u0022isItalic\u0022:false,\u0022hasNumbering\u0022:false}"
},
{
"id": "edf78020-cb71-4c75-a4d4-0e02240a193b",
"templateId": "00000000-0000-0000-0000-000000000000",
"type": "Section",
"title": "More headers",
"orderIndex": 6,
"metadataJson": "{\u0022fontSize\u0022:11,\u0022isBold\u0022:false,\u0022isItalic\u0022:false,\u0022hasNumbering\u0022:false}"
},
{
"id": "434b0aeb-d558-4cc6-91e6-9f6c1a63e511",
"templateId": "00000000-0000-0000-0000-000000000000",
"parentId": "edf78020-cb71-4c75-a4d4-0e02240a193b",
"type": "Image",
"title": "Embedded Image",
"orderIndex": 7,
"metadataJson": "{\u0022widthEmu\u0022:1928027,\u0022heightEmu\u0022:739204}"
},
{
"id": "abcebe87-d7d7-44c9-831f-384acb864723",
"templateId": "00000000-0000-0000-0000-000000000000",
"type": "Section",
"title": "See how clever word thinks it is",
"orderIndex": 8,
"metadataJson": "{\u0022fontSize\u0022:11,\u0022isBold\u0022:false,\u0022isItalic\u0022:false,\u0022hasNumbering\u0022:false}"
},
{
"id": "7e4c0259-2f20-40eb-9972-6cdad22c4da2",
"templateId": "00000000-0000-0000-0000-000000000000",
"parentId": "abcebe87-d7d7-44c9-831f-384acb864723",
"type": "Subsubsection",
"title": "111.11, 11.11, 1., A, D., (C), 1.1.1",
"orderIndex": 9,
"metadataJson": "{\u0022fontSize\u0022:11,\u0022isBold\u0022:false,\u0022isItalic\u0022:false,\u0022hasNumbering\u0022:true}"
},
{
"id": "a61fbf8e-d387-4285-84cf-9e8ee8a6d511",
"templateId": "00000000-0000-0000-0000-000000000000",
"parentId": "7e4c0259-2f20-40eb-9972-6cdad22c4da2",
"type": "List",
"title": "List Content",
"orderIndex": 10,
"metadataJson": "{\u0022listType\u0022:\u0022Numbered\u0022,\u0022items\u0022:[\u0022Football\u0022,\u0022Basketball\u0022,\u0022People\u0022]}"
},
{
"id": "aa5d5841-918c-490f-a9b9-6f1d7b839a8c",
"templateId": "00000000-0000-0000-0000-000000000000",
"parentId": "7e4c0259-2f20-40eb-9972-6cdad22c4da2",
"type": "Table",
"title": "Table Structure",
"orderIndex": 11,
"metadataJson": "{\u0022rowCount\u0022:3,\u0022colCount\u0022:3,\u0022tableData\u0022:[[\u0022Insert 1\u0022,\u0022Insert 2\u0022,\u0022Insert 3\u0022],[\u0022Dolphins\u0022,\u0022Celtics\u0022,\u0022Tiger Woods\u0022],[\u0022Lions\u0022,\u0022Heats\u0022,\u0022Mike Tyson\u0022]]}"
},
{
"id": "a5ccea1f-6327-430a-9b47-c9c62aea3d66",
"templateId": "00000000-0000-0000-0000-000000000000",
"parentId": "7e4c0259-2f20-40eb-9972-6cdad22c4da2",
"type": "Sentence",
"title": "Name",
"orderIndex": 12,
"metadataJson": "{\u0022fontSize\u0022:11,\u0022isBold\u0022:false,\u0022isItalic\u0022:true,\u0022hasNumbering\u0022:false}"
}
]
}

• Integration Instructions: how to call ParseDocxTemplate() from another project

    using TemplateParser.Core;

    var parser = new DocxParser();
    Guid myTemplateId = Guid.NewGuid();

    // Returns a ParserResult object containing a List<Node>
    var result = parser.ParseDocxTemplate("path/to/doc.docx", myTemplateId);
    //path/to/doc.docx would be the path to the document

    foreach (var node in result.Nodes)
    {
        Console.WriteLine($"[{node.Type}] {node.Title}");
    }

• Known Limitations: honest list of document patterns the parser handles poorly

*We had trouble implemening the spacing fucntion, we tried however it never actually committed to the metadata
*The font size sometimes does not update when changed, it will still say 0022:11
*We also tried doing font color and font style (times new roman, merriweather, etc), however this took a bit to implement and ended up not implementing it
*We think it would be useful for images to have a context or description box of what the image contains
*The parser requires the DocumentFormat.OpenXml library and valid .docx (XML-based) files; it does not support the legacy .doc format.
*While standard tables are captured, nested tables may have their internal metadata flattened. As well as the metadata is really messy and kinda hard to read
\*Paragraphs longer than 150 characters are automatically treated as content, even if they are bolded or large, to avoid false-positive headers.
