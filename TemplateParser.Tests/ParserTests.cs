using System.Text.Json;
using TemplateParser.Core;
using Xunit;

namespace TemplateParser.Tests;

public sealed class ParserTests
{
    private readonly DocxParser _parser = new();
    private readonly Guid _testTemplateId = Guid.NewGuid();

    // Helper to get paths for integration test files
    private string GetTestDocPath(string fileName) =>
        Path.Combine("test-documents", fileName);

    [Fact]
    public void Integration_FullPipeline_ReturnsValidJsonStructure()
    {
        // Arrange
        var path = GetTestDocPath("smoke_test.docx");

        // Act
        var result = _parser.ParseDocxTemplate(path, _testTemplateId);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result.Nodes);
        Assert.All(result.Nodes, n => Assert.Equal(_testTemplateId, n.TemplateId));
    }

    [Fact]
    public void Integration_TableDetection_CapturesMetadata()
    {
        var path = GetTestDocPath("table_sample.docx");
        var result = _parser.ParseDocxTemplate(path, _testTemplateId);

        var tableNode = result.Nodes.FirstOrDefault(n => n.Type == "Table");
        Assert.NotNull(tableNode);
        Assert.Contains("rowCount", tableNode.MetadataJson);
        Assert.Contains("tableData", tableNode.MetadataJson);
    }

    [Fact]
    public void Integration_Hierarchy_H1_IsParentOf_H2()
    {
        var path = GetTestDocPath("hierarchy_test.docx");
        var result = _parser.ParseDocxTemplate(path, _testTemplateId);

        var section = result.Nodes.First(n => n.Type == "Section");
        var subSection = result.Nodes.First(n => n.Type == "Subsection");

        Assert.Equal(section.Id, subSection.ParentId);
    }

    [Fact]
    public void Integration_ListProcessing_GroupsItemsCorrectly()
    {
        var path = GetTestDocPath("list_sample.docx");
        var result = _parser.ParseDocxTemplate(path, _testTemplateId);

        var listNode = result.Nodes.FirstOrDefault(n => n.Type == "List");
        Assert.NotNull(listNode);
        Assert.Contains("items", listNode.MetadataJson);
    }

    [Fact]
    public void Integration_EmptyDocument_ReturnsEmptyNodeList()
    {
        var path = GetTestDocPath("empty.docx");
        var result = _parser.ParseDocxTemplate(path, _testTemplateId);

        Assert.Empty(result.Nodes);
    }
}


// TODO (Week 1-6): Replace this placeholder with real tests aligned to weekly goals.
// Suggested first tests:
// - [Week 1] paragraph extraction smoke tests
// - [Week 2] heading-based hierarchy tests
// - [Week 3] table/list/image node detection tests
// - [Week 4] formatting heuristics tests for missing heading styles
// - [Week 5] integration tests that run parser through the CLI flow
// - [Week 6] refactor tests into readable groups and helper builders
//
// You may create test helpers/builders to reduce repetition (recommended by Week 6).