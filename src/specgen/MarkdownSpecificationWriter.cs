﻿using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace specgen
{
    internal class MarkdownSpecificationWriter
    {
        private static void TitleSection(XDocument spec, TextWriter writer)
        {
            var head = spec.Element("specification")?.Element("head");

            writer.WriteLine($"# {head?.Element("title")?.Value}");
            writer.WriteLine();
            writer.WriteLine(head?.Element("version")?.Value);
            writer.WriteLine();
            writer.WriteLine(head?.Element("draft")?.Value ?? head?.Element("date")?.Value);
            writer.WriteLine();
            writer.WriteLine(head?.Element("author")?.Value);
            writer.WriteLine();
            writer.WriteLine("<br/>");
            writer.WriteLine();
        }

        private static void TableOfContentsLevel(XElement level, TextWriter writer, int indent, string parentLevel)
        {
            var sections = level.Elements("section");

            var index = 1;
            foreach (var section in sections)
            {
                writer.WriteLine($"{new string(' ', indent * 2)}* [{parentLevel}{index} {section.Attribute("title").Value}](#{parentLevel}{index})");
                TableOfContentsLevel(section, writer, indent + 1, $"{parentLevel}{index}.");
                index++;
            }
        }

        private static void TableOfContents(XDocument spec, TextWriter writer)
        {
            writer.WriteLine("## Table of Contents");
            writer.WriteLine();
            TableOfContentsLevel(spec.Elements("specification").Elements("body").Single(), writer, 0, "");
        }

        public static void WriteSpecification(XDocument spec, string path)
        {
            using (var stream = new FileStream(path, FileMode.Create))
            {
                using (var writer = new StreamWriter(stream))
                {
                    TitleSection(spec, writer);
                    TableOfContents(spec, writer);
                }
            }
        }
    }
}