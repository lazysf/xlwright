using System.Text;
using XlWright.Core;
using Xunit;

namespace XlWright.McpServer.Tests;

// docs/api.md "test_run": finding the tests in tests/*.bas, the checks before Excel, and the assertion module.
public sealed class TestSourcesTests : IDisposable
{
    private static readonly Encoding ShiftJis = VbaSources.AnsiEncoding(932);

    private readonly TempRepository _repo = new();
    private readonly AppFolder _app;

    public TestSourcesTests()
    {
        _repo.AddApp("apps/sales", "Sales.xlsm");
        _app = AppFolder.Resolve(_repo.Root, "apps/sales").Folder!;
    }

    public void Dispose() => _repo.Dispose();

    [Fact]
    public void FindsSubsWithoutArgumentsThatStartWithTestInFileOrder()
    {
        Source("tests/CalcTests.bas",
            "Option Explicit\n' Public Sub TestComment()\nPublic Sub TestAdd()\nEnd Sub\nSub TestSplit ()\nEnd Sub\n" +
            "Private Sub TestHidden()\nEnd Sub\nPublic Sub TestWithArgument(ByVal x As Long)\nEnd Sub\n" +
            "Public Function TestFunction()\nEnd Function\nPublic Sub Helper()\nEnd Sub\nPublic Sub testLower()\nEnd Sub\n");
        Source("tests/Alpha.bas", "Sub TestFirst()\nEnd Sub\n");

        var plan = Prepare(null);

        Assert.Equal(["Alpha.TestFirst", "CalcTests.TestAdd", "CalcTests.TestSplit"], plan.Tests.Select(t => t.ToString()));
    }

    [Fact]
    public void ImportsTheTestModulesTheHelperClassesAndTheAssertionModule()
    {
        Source("tests/CalcTests.bas", "Public Sub TestAdd()\nEnd Sub\n");
        Source("tests/Fixture.cls", "Option Explicit\nPublic Value As Long\n");

        var plan = Prepare(null);

        Assert.Equal(["CalcTests", VbaSources.AssertName, "Fixture"], plan.Modules.Select(m => m.Name));
        Assert.Equal([VbaModuleKind.Standard, VbaModuleKind.Standard, VbaModuleKind.Class], plan.Modules.Select(m => m.Kind));
        Assert.Equal(VbaSources.AssertText(), ShiftJis.GetString(File.ReadAllBytes(plan.Modules[1].ImportPath)));
        Assert.All(plan.Modules, m => Assert.StartsWith(Path.Combine(_app.FullPath, ".xlwright", "sync"), m.ImportPath));
    }

    [Fact]
    public void RunsOnlyTheNamedTestsMatchedWithoutCase()
    {
        Source("tests/CalcTests.bas", "Public Sub TestAdd()\nEnd Sub\nPublic Sub TestSplit()\nEnd Sub\n");

        var plan = Prepare(["calctests.testsplit"]);

        Assert.Equal(["CalcTests.TestSplit"], plan.Tests.Select(t => t.ToString()));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("CalcTests.TestNone")]
    public void ReportsAMissingTestBeforeWritingAnything(string? name)
    {
        Source("tests/CalcTests.bas", "Public Sub Helper()\nEnd Sub\n" + (name is null ? "" : "Public Sub TestAdd()\nEnd Sub\n"));

        var failure = Failure(name is null ? null : [name]);

        Assert.Equal(ToolFailure.NotFound, failure.Code);
        Assert.Equal("TEST_NOT_FOUND", failure.Detail);
        if (name is not null)
        {
            Assert.Contains(name, failure.Message);
        }

        Assert.False(Directory.Exists(Path.Combine(_app.FullPath, ".xlwright", "sync")));
    }

    [Theory]
    [InlineData("tests/Calc.bas", "Calc")]
    [InlineData("tests/XlWrightAssert.bas", "XlWrightAssert")]
    [InlineData("tests/XlWrightCall.bas", "XlWrightCall")]
    [InlineData("tests/XlWright.cls", "XlWright")]
    [InlineData("tests/notes.txt", "notes.txt")]
    [InlineData("tests/sub/CalcTests.bas", "sub")]
    public void RejectsATestFileThatCannotBeImported(string file, string reason)
    {
        Source("src/modules/Calc.bas", "Option Explicit\n");
        Source("tests/CalcTests.bas", "Public Sub TestAdd()\nEnd Sub\n");
        Source(file, "Option Explicit\n");

        var failure = Failure(null);

        Assert.Equal("SOURCE_INVALID", failure.Detail);
        Assert.StartsWith(file.Contains("/sub/") ? "tests/sub" : file, failure.Data!["file"]!.GetValue<string>());
        Assert.Contains(reason, failure.Data["reason"]!.GetValue<string>());
    }

    [Fact]
    public void SkipsFilesThatStartWithADot()
    {
        Source("tests/.gitkeep", "");
        Source("tests/CalcTests.bas", "Public Sub TestAdd()\nEnd Sub\n");

        Assert.Equal(["CalcTests.TestAdd"], Prepare(null).Tests.Select(t => t.ToString()));
    }

    [Fact]
    public void TheAssertionModuleRaisesItsFailuresFromXlWrightAssert()
    {
        var text = VbaSources.AssertText();

        Assert.StartsWith($"Attribute VB_Name = \"{VbaSources.AssertName}\"\r\n", text);
        Assert.Contains("Public Sub AreEqual(", text);
        Assert.Contains("Public Sub IsTrue(", text);
        Assert.Contains("Public Sub Fail(", text);
        Assert.Contains($"\"{VbaSources.AssertName}\"", text);
    }

    private TestPlan Prepare(IReadOnlyList<string>? only)
    {
        var (plan, failure) = VbaSources.PrepareTests(_app, ShiftJis, only);
        Assert.Null(failure);
        return plan!;
    }

    private ToolFailure Failure(IReadOnlyList<string>? only)
    {
        var (plan, failure) = VbaSources.PrepareTests(_app, ShiftJis, only);
        Assert.Null(plan);
        Assert.NotNull(failure);
        return failure;
    }

    private void Source(string file, string text)
    {
        var path = Path.Combine(_app.FullPath, file);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, text, new UTF8Encoding(false));
    }
}
