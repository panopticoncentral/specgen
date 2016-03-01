using System;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Net.Mime;
using System.Runtime.InteropServices;
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
            writer.WriteLine();
            writer.WriteLine("<br/>");
            writer.WriteLine();
        }

        private static void NodeElement(XElement element, bool preserveLines, TextWriter writer)
        {
            switch (element.Name.LocalName)
            {
                case "br":
                    writer.WriteLine();
                    break;

                case "lbl":
                case "em":
                    writer.Write("**");
                    NodeText(((XText)element.Nodes().First()).Value, preserveLines, true, writer);
                    writer.Write("**");
                    break;

                case "ref":
                case "def":
                case "i":
                    writer.Write("*");
                    NodeText(((XText)element.Nodes().First()).Value, preserveLines, true, writer);
                    writer.Write("*");
                    break;

                case "emi":
                    writer.Write("**_");
                    NodeText(((XText)element.Nodes().First()).Value, preserveLines, true, writer);
                    writer.Write("_**");
                    break;

                case "c":
                    writer.Write("`");
                    NodeText(((XText)element.Nodes().First()).Value, preserveLines, true, writer);
                    writer.Write("`");
                    break;

                case "sub":
                    writer.Write("<sub>");
                    NodeText(((XText)element.Nodes().First()).Value, preserveLines, true, writer);
                    writer.Write("</sub>");
                    break;

                case "sup":
                    writer.Write("<sup>");
                    NodeText(((XText)element.Nodes().First()).Value, preserveLines, true, writer);
                    writer.Write("</sup>");
                    break;

                case "str":
                    writer.Write("~~");
                    NodeText(((XText)element.Nodes().First()).Value, preserveLines, true, writer);
                    writer.Write("~~");
                    break;
            }
        }

        private static void NodeText(string text, bool preserveLines, bool first, TextWriter writer)
        {
            var seenCharacter = false;
            var whitespacePrevious = false;

            if (preserveLines)
            {
                var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                var indent = -1;

                for (var i = 0; i < lines.Length; i++)
                {
                    var line = lines[i];

                    if (indent == -1)
                    {
                        indent = 0;

                        while (indent < line.Length && line[indent] == ' ')
                        {
                            indent++;
                        }

                        if (indent == line.Length)
                        {
                            continue;
                        }
                    }

                    if (line.Length < indent)
                    {
                        continue;
                    }

                    line = line.Substring(indent);

                    writer.Write(line);

                    if (i != lines.Length - 1)
                    {
                        indent = -1;
                        writer.WriteLine();
                    }
                }
            }
            else
            {
                foreach (var c in text)
                {
                    if (((first && !seenCharacter) || whitespacePrevious) && (c == ' ' || c == '\r' || c == '\n'))
                    {
                        continue;
                    }

                    if (c == '\r' || c == '\n')
                    {
                        writer.Write(' ');
                        whitespacePrevious = true;
                    }
                    else
                    {
                        seenCharacter = true;
                        whitespacePrevious = c == ' ';
                        writer.Write(c);
                    }
                }
            }
        }

        private static void Block(XElement block, TextWriter writer)
        {
            var preserveLines = false;

            switch (block.Name.LocalName)
            {
                case "code":
                    preserveLines = true;
                    break;
            }

            var first = true;
            foreach (var node in block.Nodes())
            {
                var element = node as XElement;
                if (element != null)
                {
                    NodeElement(element, preserveLines, writer);
                }
                else
                {
                    NodeText(((XText)node).Value, preserveLines, first, writer);
                }
                first = false;
            }

            writer.WriteLine();
        }

        private static void Section(XElement section, TextWriter writer, int level, int index, string parentLevel)
        {
            writer.WriteLine($"{new string('#', level)} <a name=\"{parentLevel}{index}\"></a>{parentLevel}{index} {section.Attribute("title").Value}");

            int subIndex = 1;
            foreach (var element in section.Elements())
            {
                writer.WriteLine();
                if (element.Name.LocalName == "section")
                {
                    Section(element, writer, level + 1, subIndex, $"{parentLevel}{index}.");
                    subIndex++;
                }
                else
                {
                    Block(element, writer);
                }
            }

            if (level == 1)
            {
                writer.WriteLine();
                writer.WriteLine("<br/>");
                writer.WriteLine();
            }
        }

        private static void Sections(XDocument spec, TextWriter writer)
        {
            int index = 1;
            foreach (var section in spec.Elements("specification").Elements("body").Single().Elements())
            {
                Section(section, writer, 1, index, "");
                index++;
            }
        }

        public static void WriteSpecification(XDocument spec, string path)
        {
            using (var stream = new FileStream(path, FileMode.Create))
            {
                using (var writer = new StreamWriter(stream))
                {
                    TitleSection(spec, writer);
                    TableOfContents(spec, writer);
                    Sections(spec, writer);
                }
            }
        }
    }
}
