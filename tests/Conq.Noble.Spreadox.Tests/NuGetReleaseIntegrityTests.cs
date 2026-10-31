using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using Conq.Noble.Spreadox;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

public class NuGetReleaseIntegrityTests
{
    private readonly ITestOutputHelper _output;

    public NuGetReleaseIntegrityTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void AssemblyVersion_And_InformationalVersion_AreDefined()
    {
        var assembly = typeof(ExcelWorkbook).Assembly;
        var version = assembly.GetName().Version;
        Assert.NotNull(version);
        Assert.True(version!.Major >= 0);

        var infoVersionAttr = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
        Assert.NotNull(infoVersionAttr);
        _output.WriteLine($"Assembly Version: {version}, Informational: {infoVersionAttr!.InformationalVersion}");
    }

    [Fact]
    public void EmbeddedFonts_ArePresentInAssemblyResources()
    {
        var assembly = typeof(ExcelWorkbook).Assembly;
        string[] resourceNames = assembly.GetManifestResourceNames();
        _output.WriteLine("Manifest resources: " + string.Join(", ", resourceNames));

        Assert.Contains(resourceNames, r => r.EndsWith("LiberationSans-Regular.ttf", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(resourceNames, r => r.EndsWith("LiberationSans-Bold.ttf", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(resourceNames, r => r.EndsWith("LiberationSans-Italic.ttf", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(resourceNames, r => r.EndsWith("LiberationSans-BoldItalic.ttf", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void VerifyDirectoryBuildProps_ReleasePackagingConfiguration()
    {
        // Locate Directory.Build.props
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
        string xmlContent = File.ReadAllText(propsPath!);
        var doc = XDocument.Parse(xmlContent);

        var packageLicense = doc.Descendants("PackageLicenseExpression").FirstOrDefault()?.Value;
        Assert.Equal("MIT", packageLicense);

        var publishRepoUrl = doc.Descendants("PublishRepositoryUrl").FirstOrDefault()?.Value;
        Assert.Equal("true", publishRepoUrl);

        var embedUntracked = doc.Descendants("EmbedUntrackedSources").FirstOrDefault()?.Value;
        Assert.Equal("true", embedUntracked);

        var includeSymbols = doc.Descendants("IncludeSymbols").FirstOrDefault()?.Value;
        Assert.Equal("true", includeSymbols);

        var symbolFormat = doc.Descendants("SymbolPackageFormat").FirstOrDefault()?.Value;
        Assert.Equal("snupkg", symbolFormat);

        _output.WriteLine("Verified Directory.Build.props packaging metadata (MIT, SourceLink, snupkg symbols).");
    }
}

