using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace Specgen
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

        private static void Term(XElement term, TextWriter writer)
        {
            switch (term.Name.LocalName)
            {
                case "nt":
                    writer.Write(term.Value);
                    break;

                case "t":
                    writer.Write($"'{term.Value}'");
                    break;

                case "meta":
                    writer.Write($"<{term.Value}>");
                    break;

                case "star":
                    Term(term.Elements().First(), writer);
                    writer.Write("*");
                    break;

                case "plus":
                    Term(term.Elements().First(), writer);
                    writer.Write("+");
                    break;

                case "opt":
                    Term(term.Elements().First(), writer);
                    writer.Write("?");
                    break;

                case "group":
                    writer.Write("(");

                    var firstSubTerm = true;

                    foreach (var subTerm in term.Elements())
                    {
                        if (!firstSubTerm)
                        {
                            writer.Write(" ");
                        }
                        firstSubTerm = false;

                        Term(subTerm, writer);
                    }

                    writer.Write(")");
                    break;

                case "range":
                    Term(term.Elements().First(), writer);
                    writer.Write("..");
                    Term(term.Elements().Skip(1).First(), writer);
                    break;
            }
        }

        private static void NodeElement(XElement element, bool preserveLines, bool first, TextWriter writer, int level)
        {
            switch (element.Name.LocalName)
            {
                case "br":
                    if (preserveLines)
                    {
                        writer.WriteLine();
                        writer.Write(new string(Enumerable.Repeat(' ', level * 2).ToArray()));
                    }
                    else
                    {
                        writer.Write("<br/>");
                    }
                    break;

                case "lbl":
                case "em":
                    writer.Write("**");
                    NodeText(((XText)element.Nodes().First()).Value, preserveLines, true, writer, level);
                    writer.Write("**");
                    break;

                case "ref":
                case "def":
                case "i":
                    writer.Write("*");
                    NodeText(((XText)element.Nodes().First()).Value, preserveLines, true, writer, level);
                    writer.Write("*");
                    break;

                case "emi":
                    writer.Write("**_");
                    NodeText(((XText)element.Nodes().First()).Value, preserveLines, true, writer, level);
                    writer.Write("_**");
                    break;

                case "c":
                    writer.Write("`");
                    NodeText(((XText)element.Nodes().First()).Value, preserveLines, true, writer, level);
                    writer.Write("`");
                    break;

                case "sub":
                    writer.Write("<sub>");
                    NodeText(((XText)element.Nodes().First()).Value, preserveLines, true, writer, level);
                    writer.Write("</sub>");
                    break;

                case "sup":
                    writer.Write("<sup>");
                    NodeText(((XText)element.Nodes().First()).Value, preserveLines, true, writer, level);
                    writer.Write("</sup>");
                    break;

                case "str":
                    writer.Write("~~");
                    NodeText(((XText)element.Nodes().First()).Value, preserveLines, true, writer, level);
                    writer.Write("~~");
                    break;

                case "token":
                case "syntax":
                    if (!first)
                    {
                        writer.WriteLine();
                    }
                    writer.WriteLine(element.Attribute("name").Value);
                    writer.Write(new string(Enumerable.Repeat(' ', (level + 1) * 2).ToArray()));

                    var firstProduction = true;

                    foreach (var production in element.Elements("production"))
                    {
                        if (firstProduction)
                        {
                            writer.Write(":");
                        }
                        else
                        {
                            writer.WriteLine();
                            writer.Write(new string(Enumerable.Repeat(' ', (level + 1) * 2).ToArray()));
                            writer.Write("|");
                        }
                        firstProduction = false;

                        foreach (var term in production.Elements())
                        {
                            writer.Write(" ");
                            Term(term, writer);
                        }
                    }

                    writer.WriteLine();
                    writer.Write(new string(Enumerable.Repeat(' ', (level + 1) * 2).ToArray()));
                    writer.WriteLine(";");
                    break;
            }
        }

        private static void NodeText(string text, bool preserveLines, bool first, TextWriter writer, int level)
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
                        writer.Write(new string(Enumerable.Repeat(' ', level * 2).ToArray()));
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

        private static void Block(XElement block, TextWriter writer, int level = 0)
        {
            var preserveLines = false;

            if (level > 0)
            {
                switch (block.Name.LocalName)
                {
                    case "bulletedText":
                        writer.Write(new string(Enumerable.Repeat(' ', (level - 1) * 2).ToArray()));
                        writer.Write("* ");
                        break;

                    case "numberedText":
                        writer.Write(new string(Enumerable.Repeat(' ', (level - 1) * 2).ToArray()));
                        writer.Write("1. ");
                        break;

                    default:
                        writer.Write(new string(Enumerable.Repeat(' ', level * 2).ToArray()));
                        break;
                }
            }

            switch (block.Name.LocalName)
            {
                case "alert":
                case "annotation":
                    writer.Write("> ");
                    break;

                case "code":
                    writer.WriteLine("```");
                    writer.Write(new string(Enumerable.Repeat(' ', level * 2).ToArray()));
                    preserveLines = true;
                    break;

                case "bulletedList":
                case "numberedList":
                    var firstItem = true;
                    foreach (var item in block.Elements())
                    {
                        if (firstItem)
                        {
                            firstItem = false;
                        }
                        else
                        {
                            writer.WriteLine();
                        }
                        Block(item, writer, level + 1);
                    }
                    return;

                case "grammar":
                    writer.WriteLine("```antlr");
                    writer.Write(new string(Enumerable.Repeat(' ', level * 2).ToArray()));
                    break;
            }

            var first = true;
            foreach (var node in block.Nodes())
            {
                if (node is XElement element)
                {
                    NodeElement(element, preserveLines, first, writer, level);
                }
                else
                {
                    NodeText(((XText)node).Value, preserveLines, first, writer, level);
                }
                first = false;
            }

            switch (block.Name.LocalName)
            {
                case "grammar":
                case "code":
                    writer.Write("```");
                    break;
            }

            writer.WriteLine();
        }

        private static void Section(XElement section, TextWriter writer, int level, int index, string parentLevel)
        {
            writer.WriteLine($"{new string('#', level)} <a name=\"{parentLevel}{index}\"></a>{parentLevel}{index} {section.Attribute("title").Value}");

            var subIndex = 1;
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
            var index = 1;
            foreach (var section in spec.Elements("specification").Elements("body").Single().Elements())
            {
                Section(section, writer, 1, index, "");
                index++;
            }
        }

        public static void WriteSpecification(XDocument spec, string path)
        {
            using var stream = new FileStream(path, FileMode.Create);
            using var writer = new StreamWriter(stream);
            TitleSection(spec, writer);
            TableOfContents(spec, writer);
            Sections(spec, writer);
        }
    }
}
