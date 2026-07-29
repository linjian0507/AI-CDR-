using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace AIVectorCore
{
    public sealed class SvgLayer
    {
        public string Name { get; set; } = "";
        public string Id { get; set; } = "";
        public int Index { get; set; }
        public string Svg { get; set; } = "";
    }

    public static class SvgLayers
    {
        public static IReadOnlyList<SvgLayer> Parse(string svg)
        {
            if (string.IsNullOrWhiteSpace(svg)) return new List<SvgLayer>();
            try
            {
                var document = XDocument.Parse(svg, LoadOptions.PreserveWhitespace);
                var root = document.Root;
                if (root == null) return new List<SvgLayer>();
                var result = new List<SvgLayer>();
                var index = 0;
                foreach (var child in root.Elements())
                {
                    if (!string.Equals(child.Name.LocalName, "g", StringComparison.OrdinalIgnoreCase)) continue;
                    var name = (string)child.Attribute("data-name") ?? (string)child.Attribute("id") ?? "";
                    result.Add(new SvgLayer
                    {
                        Name = name,
                        Id = (string)child.Attribute("id") ?? "",
                        Index = index++,
                        Svg = child.ToString(SaveOptions.DisableFormatting)
                    });
                }
                return result;
            }
            catch
            {
                return ParseFallback(svg);
            }
        }

        private static IReadOnlyList<SvgLayer> ParseFallback(string svg)
        {
            var result = new List<SvgLayer>();
            var regex = new Regex(@"<(/?)(g|svg|defs)\b([^>]*?)(/?)>", RegexOptions.IgnoreCase);
            var depth = 0;
            var inSvg = false;
            var inDefs = false;
            Match current = null;
            var index = 0;

            foreach (Match match in regex.Matches(svg))
            {
                var closing = match.Groups[1].Value == "/";
                var tag = match.Groups[2].Value.ToLowerInvariant();
                var attrs = match.Groups[3].Value ?? "";
                var selfClosing = match.Groups[4].Value == "/";
                if (tag == "svg")
                {
                    inSvg = !closing;
                    continue;
                }
                if (!inSvg) continue;
                if (tag == "defs")
                {
                    if (!selfClosing) inDefs = !closing;
                    continue;
                }
                if (inDefs || tag != "g") continue;

                if (!closing)
                {
                    if (depth == 0)
                    {
                        current = match;
                        if (selfClosing)
                        {
                            result.Add(BuildFallbackLayer(svg, current, match.Index + match.Length, index++));
                            current = null;
                        }
                    }
                    if (!selfClosing) depth++;
                }
                else
                {
                    depth--;
                    if (depth == 0 && current != null)
                    {
                        result.Add(BuildFallbackLayer(svg, current, match.Index + match.Length, index++));
                        current = null;
                    }
                    if (depth < 0) depth = 0;
                }
            }
            return result;
        }

        private static SvgLayer BuildFallbackLayer(string svg, Match start, int end, int index)
        {
            var attrs = start.Groups[3].Value ?? "";
            var name = AttributeValue(attrs, "data-name");
            if (string.IsNullOrEmpty(name)) name = AttributeValue(attrs, "id");
            return new SvgLayer
            {
                Name = name,
                Id = AttributeValue(attrs, "id"),
                Index = index,
                Svg = svg.Substring(start.Index, Math.Max(0, Math.Min(end, svg.Length) - start.Index))
            };
        }

        private static string AttributeValue(string attrs, string name)
        {
            var match = Regex.Match(attrs, @"\b" + Regex.Escape(name) + @"\s*=\s*""([^""]*)""", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value : "";
        }
    }
}
