using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using Conq.Noble.Spreadox;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

public class ReproducibleBuildTests
{
    private readonly ITestOutputHelper _output;

    public ReproducibleBuildTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void BuildConfiguration_EnforcesDeterministicAndSourceLinkProperties()
    {
        string currentDir = AppDomain.CurrentDomain.BaseDirectory;
        DirectoryInfo? dir = new DirectoryInfo(currentDir);
        string? propsPath = null;

        while (dir != null)
        {
            string candidate = Path.Combine(dir.FullName, "Directory.Build.props");
            if (File.Exists(candidate))
            {
                propsPath = candidate;
                break;
            }
            dir = dir.Parent;
        }

        Assert.NotNull(propsPath);
        var doc = XDocument.Parse(File.ReadAllText(propsPath!));

        // 1. Deterministic build setting
        var deterministic = doc.Descendants("Deterministic").FirstOrDefault()?.Value;
        Assert.Equal("true", deterministic);

        // 2. SourceLink and symbol properties
        var publishRepoUrl = doc.Descendants("PublishRepositoryUrl").FirstOrDefault()?.Value;
        Assert.Equal("true", publishRepoUrl);

        var embedUntracked = doc.Descendants("EmbedUntrackedSources").FirstOrDefault()?.Value;
        Assert.Equal("true", embedUntracked);

        var includeSymbols = doc.Descendants("IncludeSymbols").FirstOrDefault()?.Value;
        Assert.Equal("true", includeSymbols);

        var symbolFormat = doc.Descendants("SymbolPackageFormat").FirstOrDefault()?.Value;
        Assert.Equal("snupkg", symbolFormat);

        _output.WriteLine("Verified: Deterministic and SourceLink settings are strictly configured in Directory.Build.props.");
    }

    [Fact]
    public void AssemblyAttributes_ContainExpectedReproducibilityMetadata()
    {
        var assembly = typeof(ExcelWorkbook).Assembly;

        // Assembly should have InformationalVersionAttribute
        var infoVersion = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
        Assert.NotNull(infoVersion);
        Assert.False(string.IsNullOrWhiteSpace(infoVersion!.InformationalVersion));

        // Assembly should have Company, Product, Copyright metadata
        var product = assembly.GetCustomAttribute<AssemblyProductAttribute>()?.Product;
        Assert.Equal("Spreadox", product);

        var company = assembly.GetCustomAttribute<AssemblyCompanyAttribute>()?.Company;
        Assert.Equal("Conq.Noble", company);

        _output.WriteLine($"Assembly InformationalVersion: {infoVersion.InformationalVersion}, Product: {product}, Company: {company}");
    }
}
