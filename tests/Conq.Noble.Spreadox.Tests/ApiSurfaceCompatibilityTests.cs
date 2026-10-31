using System;
using System.Linq;
using System.Reflection;
using Conq.Noble.Spreadox;
using Xunit;
using Xunit.Abstractions;

namespace Conq.Noble.Spreadox.Tests;

/// <summary>
/// Group 12: API Surface Compatibility & Dependency Isolation Tests (P1)
/// Ensures zero vendor type leaks and preserves public binary compatibility.
/// </summary>
public class ApiSurfaceCompatibilityTests
{
    private readonly ITestOutputHelper _output;

    public ApiSurfaceCompatibilityTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void VendorTypes_NeverLeakIntoPublicApiSurface()
    {
        var assembly = typeof(ExcelWorkbook).Assembly;
        var publicTypes = assembly.GetExportedTypes();

        foreach (Type type in publicTypes)
        {
            // 1. Check base type and interfaces
            if (type.BaseType != null)
            {
                AssertDoesNotReferenceVendor(type.BaseType, type);
            }

            foreach (Type iface in type.GetInterfaces())
            {
                AssertDoesNotReferenceVendor(iface, type);
            }

            // 2. Check public constructors
            foreach (ConstructorInfo ctor in type.GetConstructors(BindingFlags.Public | BindingFlags.Instance))
            {
                foreach (ParameterInfo p in ctor.GetParameters())
                {
                    AssertDoesNotReferenceVendor(p.ParameterType, type, ctor.Name);
                }
            }

            // 3. Check public methods
            foreach (MethodInfo method in type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly))
            {
                AssertDoesNotReferenceVendor(method.ReturnType, type, method.Name);
                foreach (ParameterInfo p in method.GetParameters())
                {
                    AssertDoesNotReferenceVendor(p.ParameterType, type, method.Name);
                }
            }

            // 4. Check public properties
            foreach (PropertyInfo prop in type.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly))
            {
                AssertDoesNotReferenceVendor(prop.PropertyType, type, prop.Name);
            }
        }

        _output.WriteLine($"Verified {publicTypes.Length} public types: zero OpenXML or PdfSharpCore vendor types leaked.");
    }

    private static void AssertDoesNotReferenceVendor(Type t, Type owningType, string? memberName = null)
    {
        Type checkType = t.IsGenericType ? t.GetGenericTypeDefinition() : t;
        string? ns = checkType.Namespace;

        bool isVendor = ns != null && (
            ns.StartsWith("DocumentFormat.OpenXml", StringComparison.OrdinalIgnoreCase) ||
            ns.StartsWith("PdfSharpCore", StringComparison.OrdinalIgnoreCase));

        Assert.False(isVendor, $"Vendor type '{checkType.FullName}' is exposed in public surface of '{owningType.FullName}.{memberName}'");
    }

    [Fact]
    public void KeyPublicTypes_ExistWithExpectedContract()
    {
        var assembly = typeof(ExcelWorkbook).Assembly;
        string[] expectedTypeNames =
        {
            "Conq.Noble.Spreadox.ExcelWorkbook",
            "Conq.Noble.Spreadox.ExcelStreamingWorkbook",
            "Conq.Noble.Spreadox.ExcelWorksheet",
            "Conq.Noble.Spreadox.ExcelCell",
            "Conq.Noble.Spreadox.ExcelChart",
            "Conq.Noble.Spreadox.ExcelRange",
            "Conq.Noble.Spreadox.CellAddress",
            "Conq.Noble.Spreadox.PdfReport",
            "Conq.Noble.Spreadox.PdfFontRegistry",
            "Conq.Noble.Spreadox.NumberFormats"
        };

        foreach (string name in expectedTypeNames)
        {
            Type? t = assembly.GetType(name, throwOnError: false);
            Assert.NotNull(t);
            Assert.True(t!.IsPublic, $"Type '{name}' should be public.");
        }

        _output.WriteLine("All key public abstractions exist and conform to public surface contract.");
    }
}
