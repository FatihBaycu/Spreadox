using System;
using System.Collections.Concurrent;
using System.IO;
using System.Reflection;
using PdfSharpCore.Fonts;

namespace Conq.Noble.Spreadox.Internal
{
    internal sealed class PdfFontResolver : IFontResolver
    {
        internal const string FamilyName = "Liberation Sans";

        private const string FaceRegular = "LiberationSans#R";
        private const string FaceBold = "LiberationSans#B";
        private const string FaceItalic = "LiberationSans#I";
        private const string FaceBoldItalic = "LiberationSans#BI";

        private static readonly object Gate = new object();
        private static bool _installed;

        // Custom font storage
        private static readonly ConcurrentDictionary<string, byte[]> CustomFonts
            = new ConcurrentDictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

        // Track which family names have been registered
        private static readonly ConcurrentDictionary<string, bool> RegisteredFamilies
            = new ConcurrentDictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

        public string DefaultFontName => FamilyName;

        internal static void EnsureInstalled()
        {
            if (_installed)
            {
                return;
            }

            lock (Gate)
            {
                if (_installed)
                {
                    return;
                }

                GlobalFontSettings.FontResolver = new PdfFontResolver();
                _installed = true;
            }
        }

        /// <summary>
        /// Register a custom TrueType font from raw bytes. Call before generating any PDF.
        /// </summary>
        public static void RegisterFont(string familyName, byte[] ttfBytes, bool isBold = false, bool isItalic = false)
        {
            if (string.IsNullOrWhiteSpace(familyName))
            {
                throw new ArgumentException("Family name is required.", nameof(familyName));
            }

            if (ttfBytes is null || ttfBytes.Length == 0)
            {
                throw new ArgumentException("Font bytes are required.", nameof(ttfBytes));
            }

            string faceName = BuildFaceName(familyName, isBold, isItalic);
            CustomFonts[faceName] = (byte[])ttfBytes.Clone();
            RegisteredFamilies[familyName] = true;
        }

        /// <summary>
        /// Register a custom TrueType font from a file path. Call before generating any PDF.
        /// </summary>
        public static void RegisterFont(string familyName, string ttfFilePath, bool isBold = false, bool isItalic = false)
        {
            if (string.IsNullOrWhiteSpace(ttfFilePath))
            {
                throw new ArgumentException("File path is required.", nameof(ttfFilePath));
            }

            RegisterFont(familyName, File.ReadAllBytes(ttfFilePath), isBold, isItalic);
        }

        /// <summary>
        /// Check whether a custom font family has been registered.
        /// </summary>
        internal static bool IsFamilyRegistered(string familyName)
        {
            return !string.IsNullOrWhiteSpace(familyName)
                && RegisteredFamilies.ContainsKey(familyName);
        }

        private static string BuildFaceName(string familyName, bool isBold, bool isItalic)
        {
            string suffix = (isBold, isItalic) switch
            {
                (true, true) => "#BI",
                (true, false) => "#B",
                (false, true) => "#I",
                _ => "#R"
            };
            return familyName + suffix;
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            // Check custom fonts first
            if (!string.IsNullOrEmpty(familyName) && RegisteredFamilies.ContainsKey(familyName))
            {
                string face = BuildFaceName(familyName, isBold, isItalic);
                // If exact match not found, fall back to regular of that family
                if (!CustomFonts.ContainsKey(face))
                {
                    face = BuildFaceName(familyName, false, false);
                }
                if (CustomFonts.ContainsKey(face))
                {
                    return new FontResolverInfo(face);
                }
            }

            // Default Liberation Sans
            if (isBold && isItalic)
            {
                return new FontResolverInfo(FaceBoldItalic);
            }

            if (isBold)
            {
                return new FontResolverInfo(FaceBold);
            }

            if (isItalic)
            {
                return new FontResolverInfo(FaceItalic);
            }

            return new FontResolverInfo(FaceRegular);
        }

        public byte[] GetFont(string faceName)
        {
            // Check custom fonts first
            if (CustomFonts.TryGetValue(faceName, out byte[]? customBytes))
            {
                return customBytes;
            }

            // Default Liberation Sans
            string resource = faceName switch
            {
                FaceBold => "LiberationSans-Bold.ttf",
                FaceItalic => "LiberationSans-Italic.ttf",
                FaceBoldItalic => "LiberationSans-BoldItalic.ttf",
                _ => "LiberationSans-Regular.ttf"
            };

            return LoadEmbedded(resource);
        }

        private static readonly ConcurrentDictionary<string, byte[]> EmbeddedFontCache =
            new ConcurrentDictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

        private static byte[] LoadEmbedded(string fileName)
        {
            return EmbeddedFontCache.GetOrAdd(fileName, name =>
            {
                Assembly assembly = typeof(PdfFontResolver).Assembly;
                string resourceName = "Conq.Noble.Spreadox.Fonts." + name;
                Stream? stream = assembly.GetManifestResourceStream(resourceName);
                if (stream is null)
                {
                    throw new InvalidOperationException("Embedded font not found: " + resourceName);
                }

                using (stream)
                using (var copy = new MemoryStream())
                {
                    stream.CopyTo(copy);
                    return copy.ToArray();
                }
            });
        }
    }
}
