using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace specgen
{
    internal class Program
    {
        private static readonly XNamespace Pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        // ReSharper disable once InconsistentNaming
        private static readonly XNamespace rs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        // ReSharper disable once InconsistentNaming
        private static readonly XNamespace prs = "http://schemas.openxmlformats.org/package/2006/relationships";
        // ReSharper disable once InconsistentNaming
        private static readonly XNamespace ws = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        private static readonly XNamespace o = "urn:schemas-microsoft-com:office:office";
        private static readonly XNamespace v = "urn:schemas-microsoft-com:vml";

        private struct Relationship
        {
            public string Id;
            public string Type;
            public string Target;
        }

        private struct FontSignature
        {
            public string Usb0;
            public string Usb1;
            public string Usb2;
            public string Usb3;
            public string Csb0;
            public string Csb1;
        }

        private struct Font
        {
            public string Name;
            public string Panose1;
            public string Charset;
            public string Family;
            public string Pitch;
            public FontSignature Sig;
        }

        private static XElement CreateRelationships(IEnumerable<Relationship> relationships)
        {
            return new XElement(prs + "Relationships",
                from r in relationships
                select new XElement(prs + "Relationship",
                    new XAttribute("Id", r.Id),
                    new XAttribute("Type", $"{rs.NamespaceName}/{r.Type}"),
                    new XAttribute("Target", r.Target)));
        }

        private static XElement Part(string name, string contentType, string padding, XElement data)
        {
            return new XElement(Pkg + "part",
                new XAttribute(Pkg + "name", name),
                new XAttribute(Pkg + "contentType", contentType),
                (padding != null) ? new XAttribute(Pkg + "padding", padding) : null,
                new XElement(Pkg + "xmlData",
                    data));
        }

        private static XElement PackageRelationships()
        {
            return Part(
                "/_rels/.rels",
                "application/vnd.openxmlformats-package.relationships+xml",
                "512",
                CreateRelationships(new List<Relationship>
                {
                    new Relationship {Id = "rId1", Type = "officeDocument", Target = "word/document.xml"}
                }));
        }

        private static XElement DocumentRelationships()
        {
            return Part(
                "/word/_rels/document.xml.rels",
                "application/vnd.openxmlformats-package.relationships+xml",
                "256",
                CreateRelationships(new List<Relationship>
                {
                    new Relationship {Id = "rId1", Type = "fontTable", Target = "fontTable.xml"},
                    new Relationship {Id = "rId2", Type = "footer", Target = "footer1.xml"},
                    new Relationship {Id = "rId3", Type = "footer", Target = "footer2.xml"},
                    new Relationship {Id = "rId4", Type = "footer", Target = "footer3.xml"},
                    new Relationship {Id = "rId5", Type = "footer", Target = "footer4.xml"},
                    new Relationship {Id = "rId6", Type = "footer", Target = "footer5.xml"},
                    new Relationship {Id = "rId7", Type = "header", Target = "header1.xml"},
                    new Relationship {Id = "rId8", Type = "header", Target = "header2.xml"},
                    new Relationship {Id = "rId9", Type = "header", Target = "header3.xml"},
                    new Relationship {Id = "rId10", Type = "header", Target = "header4.xml"},
                    new Relationship {Id = "rId11", Type = "header", Target = "header5.xml"},
                    new Relationship {Id = "rId12", Type = "header", Target = "header6.xml"},
                    new Relationship {Id = "rId13", Type = "numbering", Target = "numbering.xml"},
                    new Relationship {Id = "rId14", Type = "settings", Target = "settings.xml"},
                    new Relationship {Id = "rId15", Type = "styles", Target = "styles.xml"},
                    new Relationship {Id = "rId16", Type = "webSettings", Target = "webSettings.xml"}
                }));
        }

        private static XElement KeyValue(XNamespace ns, string key, string value)
        {
            return new XElement(ns + key,
                new XAttribute(ns + "val", value));
        }

        private static XElement KeyValue(string key, string value)
        {
            return KeyValue(ws, key, value);
        }

        private static XElement KeyValue(string key, string valueName, string value)
        {
            return new XElement(ws + key,
                new XAttribute(ws + valueName, value));
        }

        private static XElement Fonts(IEnumerable<Font> fonts)
        {
            return new XElement(ws + "fonts",
                new XAttribute(XNamespace.Xmlns + "w", ws.NamespaceName),
                from f in fonts
                select new XElement(ws + "font",
                    new XAttribute(ws + "name", f.Name),
                    KeyValue("panose1", f.Panose1),
                    KeyValue("charset", f.Charset),
                    KeyValue("family", f.Family),
                    KeyValue("pitch", f.Pitch),
                    new XElement(ws + "sig",
                        new XAttribute(ws + "usb0", f.Sig.Usb0),
                        new XAttribute(ws + "usb1", f.Sig.Usb1),
                        new XAttribute(ws + "usb2", f.Sig.Usb2),
                        new XAttribute(ws + "usb3", f.Sig.Usb3),
                        new XAttribute(ws + "csb0", f.Sig.Csb0),
                        new XAttribute(ws + "csb1", f.Sig.Csb1)
                        )
                    ));
        }

        private static XElement FontTable()
        {
            return Part(
                "/word/fontTable.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
                null,
                Fonts(new List<Font>
                {
                    new Font
                    {
                        Name ="Arial",
                        Panose1="020B0604020202020204",
                        Charset="00",
                        Family="swiss",
                        Pitch="variable",
                        Sig = new FontSignature { Usb0="E0002AFF", Usb1="C0007843", Usb2="00000009", Usb3="00000000", Csb0="000001FF", Csb1="00000000"}
                    },
                    new Font
                    {
                        Name ="Calibri",
                        Panose1="020F0502020204030204",
                        Charset="00",
                        Family="swiss",
                        Pitch="variable",
                        Sig = new FontSignature { Usb0="E10002FF", Usb1="4000ACFF", Usb2="00000009", Usb3="00000000", Csb0="0000019F", Csb1="00000000"}
                    },
                    new Font
                    {
                        Name ="Cambria",
                        Panose1="02040503050406030204",
                        Charset="00",
                        Family="roman",
                        Pitch="variable",
                        Sig = new FontSignature { Usb0="A00002EF", Usb1="4000004B", Usb2="00000000", Usb3="00000000", Csb0="0000019F", Csb1="00000000"}
                    },
                    new Font
                    {
                        Name ="Cambria Math",
                        Panose1="02040503050406030204",
                        Charset="00",
                        Family="roman",
                        Pitch="variable",
                        Sig = new FontSignature { Usb0="A00002EF", Usb1="420020EB", Usb2="00000000", Usb3="00000000", Csb0="0000019F", Csb1="00000000"}
                    },
                    new Font
                    {
                        Name ="Consolas",
                        Panose1="020B0609020204030204",
                        Charset="00",
                        Family="modern",
                        Pitch="fixed",
                        Sig = new FontSignature { Usb0="E10002FF", Usb1="4000FCFF", Usb2="00000009", Usb3="00000000", Csb0="0000019F", Csb1="00000000"}
                    },
                    new Font
                    {
                        Name ="Courier New",
                        Panose1="02070309020205020404",
                        Charset="00",
                        Family="modern",
                        Pitch="fixed",
                        Sig = new FontSignature { Usb0="E0002AFF", Usb1="C0007843", Usb2="00000009", Usb3="00000000", Csb0="000001FF", Csb1="00000000"}
                    },
                    new Font
                    {
                        Name ="Symbol",
                        Panose1="05050102010706020507",
                        Charset="02",
                        Family="roman",
                        Pitch="variable",
                        Sig = new FontSignature { Usb0="00000000", Usb1="10000000", Usb2="00000000", Usb3="00000000", Csb0="80000000", Csb1="00000000"}
                    },
                    new Font
                    {
                        Name ="Tahoma",
                        Panose1="020B0604030504040204",
                        Charset="00",
                        Family="swiss",
                        Pitch="variable",
                        Sig = new FontSignature { Usb0="E1002EFF", Usb1="C000605B", Usb2="00000029", Usb3="00000000", Csb0="000101FF", Csb1="00000000"}
                    },
                    new Font
                    {
                        Name ="Times New Roman",
                        Panose1="02020603050405020304",
                        Charset="00",
                        Family="roman",
                        Pitch="variable",
                        Sig = new FontSignature { Usb0="E0002AFF", Usb1="C0007841", Usb2="00000009", Usb3="00000000", Csb0="000001FF", Csb1="00000000"}
                    }
                }));
        }

        static IEnumerable<object> TitleSection(XDocument spec)
        {
            var head = spec.Element("specification")?.Element("head");

            for (var index = 0; index < 8; index++)
            {
                yield return Para("Text");
            }

            yield return Para("Title",
                Run(
                    Text(head?.Element("title")?.Value, true)));

            for (var index = 0; index < 7; index++)
            {
                yield return Para("Text");
            }

            yield return Para("Subtitle",
                Run(
                    Text(head?.Element("version")?.Value, true)));

            yield return Para("Subtitle",
                Run(
                    Text(head?.Element("draft")?.Value ?? head?.Element("date")?.Value, true)));

            for (var index = 0; index < 3; index++)
            {
                yield return Para("Text");
            }

            yield return Para("Author",
                Run(
                    Text(head?.Element("author")?.Value, true)));

            yield return Para(
                ParaProperties(
                    SectionProperties(
                        new XElement(ws + "footerReference",
                            new XAttribute(ws + "type", "even"),
                            new XAttribute(rs + "id", "rId2")),
                        new XElement(ws + "footerReference",
                            new XAttribute(ws + "type", "default"),
                            new XAttribute(rs + "id", "rId3")),
                        new XElement(ws + "pgSz",
                            new XAttribute(ws + "w", "12240"),
                            new XAttribute(ws + "h", "15840")),
                        new XElement(ws + "pgMar",
                            new XAttribute(ws + "top", "1440"),
                            new XAttribute(ws + "right", "1660"),
                            new XAttribute(ws + "bottom", "1440"),
                            new XAttribute(ws + "left", "1660"),
                            new XAttribute(ws + "header", "1020"),
                            new XAttribute(ws + "footer", "1020"),
                            new XAttribute(ws + "gutter", "0")),
                        new XElement(ws + "cols",
                            new XAttribute(ws + "space", "720")),
                        new XElement(ws + "docGrid",
                            new XAttribute(ws + "linePitch", "360")))));
        }

        static IEnumerable<object> TocSection()
        {
            yield return Para("TOCHeading",
                Run(
                    new XElement(ws + "lastRenderedPageBreak"),
                    Text("Table of Contents")));

            yield return Para("Text");

            yield return Para(
                Field(" TOC \\o \"3-9\" \\h \\z \\t \"Heading 1,1,Heading 2,2\" "));

            yield return Para(
                ParaProperties(
                    SectionProperties(
                        new XElement(ws + "footerReference",
                            new XAttribute(ws + "type", "even"),
                            new XAttribute(rs + "id", "rId4")),
                        new XElement(ws + "footerReference",
                            new XAttribute(ws + "type", "default"),
                            new XAttribute(rs + "id", "rId5")),
                        new XElement(ws + "footerReference",
                            new XAttribute(ws + "type", "first"),
                            new XAttribute(rs + "id", "rId6")),
                        new XElement(ws + "headerReference",
                            new XAttribute(ws + "type", "even"),
                            new XAttribute(rs + "id", "rId7")),
                        new XElement(ws + "headerReference",
                            new XAttribute(ws + "type", "default"),
                            new XAttribute(rs + "id", "rId8")),
                        new XElement(ws + "headerReference",
                            new XAttribute(ws + "type", "first"),
                            new XAttribute(rs + "id", "rId9")),
                        KeyValue("type", "oddPage"),
                        new XElement(ws + "pgSz",
                            new XAttribute(ws + "w", "12240"),
                            new XAttribute(ws + "h", "15840")),
                        new XElement(ws + "pgMar",
                            new XAttribute(ws + "top", "1440"),
                            new XAttribute(ws + "right", "1152"),
                            new XAttribute(ws + "bottom", "1440"),
                            new XAttribute(ws + "left", "1152"),
                            new XAttribute(ws + "header", "1022"),
                            new XAttribute(ws + "footer", "1022"),
                            new XAttribute(ws + "gutter", "0")),
                        new XElement(ws + "pgNumType",
                            new XAttribute(ws + "fmt", "lowerRoman"),
                            new XAttribute(ws + "start", "1")),
                        new XElement(ws + "cols",
                            new XAttribute(ws + "space", "720")),
                        new XElement(ws + "titlePg"),
                        new XElement(ws + "docGrid",
                            new XAttribute(ws + "linePitch", "360")))));
        }

        static IEnumerable<object> Section(XElement section, int level, bool first)
        {
            yield return Para($"Heading{level}",
                Run(
                    level == 1 ? new XElement(ws + "lastRenderedPageBreak") : null,
                    Text(section.Attribute("title").Value)));



            if (level == 1)
            {
                yield return Para(
                    ParaProperties(
                        SectionProperties(
                            first ? new XElement(ws + "headerReference",
                                new XAttribute(ws + "type", "even"),
                                new XAttribute(rs + "id", "rId10")) : null,
                            first ? new XElement(ws + "headerReference",
                                new XAttribute(ws + "type", "default"),
                                new XAttribute(rs + "id", "rId11")) : null,
                            first ? new XElement(ws + "headerReference",
                                new XAttribute(ws + "type", "first"),
                                new XAttribute(rs + "id", "rId12")) : null,
                            KeyValue("type", "oddPage"),
                            new XElement(ws + "pgSz",
                                new XAttribute(ws + "w", "12240"),
                                new XAttribute(ws + "h", "15840")),
                            new XElement(ws + "pgMar",
                                new XAttribute(ws + "top", "1440"),
                                new XAttribute(ws + "right", "1152"),
                                new XAttribute(ws + "bottom", "1440"),
                                new XAttribute(ws + "left", "1152"),
                                new XAttribute(ws + "header", "1022"),
                                new XAttribute(ws + "footer", "1022"),
                                new XAttribute(ws + "gutter", "0")),
                            first ? new XElement(ws + "pgNumType",
                                new XAttribute(ws + "start", "1")) : null,
                            new XElement(ws + "cols",
                                new XAttribute(ws + "space", "720")),
                            new XElement(ws + "titlePg"),
                            new XElement(ws + "docGrid",
                                new XAttribute(ws + "linePitch", "360")))));
            }
        } 

        static IEnumerable<object> DocumentSections(XDocument spec)
        {
            var sections = spec.Descendants("specification").Descendants("body").Descendants("section");

            yield return TitleSection(spec);
            yield return TocSection();

            var first = true;
            foreach (var section in sections)
            {
                yield return Section(section, 1, first);
                first = false;
            }
        }

        static XElement Document(XDocument spec)
        {
            return Part("/word/document.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                null,
                new XElement(ws + "document",
                    new XAttribute(XNamespace.Xmlns + "w", ws.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "r", rs.NamespaceName),
                    new XElement(ws + "body",
                        DocumentSections(spec))));
        }

        static XElement Para(string style, params object[] elements)
        {
            return Para(
                ParaProperties(
                    KeyValue("pStyle", style)),
                elements);
        }

        static XElement Para(params object[] elements)
        {
            return new XElement(ws + "p",
                elements);
        }

        static XElement Run(params object[] elements)
        {
            return new XElement(ws + "r",
                elements);
        }

        static XElement Text(string text, bool preserve = false)
        {
            return new XElement(ws + "t",
                preserve ? new XAttribute(XNamespace.Xml + "space", "preserve") : null,
                text);
        }

        static XElement Break()
        {
            return new XElement(ws + "br");
        }

        static XElement Tab()
        {
            return new XElement(ws + "tab");
        }

        static XElement SectionProperties(params object[] elements)
        {
            return new XElement(ws + "sectPr",
                elements);
        }

        static XElement ParaProperties(params object[] elements)
        {
            return new XElement(ws + "pPr",
                elements);
        }

        static XElement RunProperties(params object[] elements)
        {
            return new XElement(ws + "rPr",
                elements);
        }

        static IEnumerable<XElement> Field(string value, params object[] properties)
        {
            return new List<XElement>
            {
                Run(
                    properties != null && properties.Length > 0 ? RunProperties(properties) : null,
                    new XElement(ws + "fldChar",
                        new XAttribute(ws + "fldCharType", "begin"))),
                Run(
                    properties != null && properties.Length > 0 ? RunProperties(properties) : null,
                    new XElement(ws + "instrText",
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        value)),
                Run(
                    properties != null && properties.Length > 0 ? RunProperties(properties) : null,
                    new XElement(ws + "fldChar",
                        new XAttribute(ws + "fldCharType", "separate"))),
                Run(
                    properties != null && properties.Length > 0 ? RunProperties(properties) : null,
                    new XElement(ws + "fldChar",
                        new XAttribute(ws + "fldCharType", "end")))
            };
        }

        static XElement Tabs(params Tuple<string, string>[] tabs)
        {
            return new XElement(ws + "tabs",
                from t in tabs
                select new XElement(ws + "tab",
                    new XAttribute(ws + "val", t.Item1),
                    new XAttribute(ws + "pos", t.Item2)));
        }

        static XElement Tabs(params Tuple<string, string, string>[] tabs)
        {
            return new XElement(ws + "tabs",
                from t in tabs
                select new XElement(ws + "tab",
                    new XAttribute(ws + "val", t.Item1),
                    new XAttribute(ws + "pos", t.Item2),
                    new XAttribute(ws + "leader", t.Item3)));
        }

        static XElement StandardTabs()
        {
            return Tabs(
                new Tuple<string, string>("clear", "4320"),
                new Tuple<string, string>("clear", "8640"),
                new Tuple<string, string>("right", "9936"));
        }

        static XElement Footer(string name, params object[] elements)
        {
            return Part($"/word/{name}.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
                null,
                new XElement(ws + "ftr",
                    new XAttribute(XNamespace.Xmlns + "w", ws.NamespaceName),
                    Para(
                        elements)));
        }

        static IEnumerable<XElement> Footers()
        {
            return new List<XElement>
            {
                Footer("footer1",
                    ParaProperties(
                        KeyValue("pStyle", "Footer"),
                        KeyValue("ind", "right", "360"))),
                Footer("footer2",
                    ParaProperties(
                        KeyValue("pStyle", "Footer"),
                        KeyValue("jc", "center")),
                    Run(
                        Text("Copyright © ", true)),
                    Field(" DATE  \\@ \"yyyy\"  \\* MERGEFORMAT "),
                    Run(
                        Text(". All Rights Reserved."),
                        Break())),
                Footer("footer3",
                    ParaProperties(
                        KeyValue("pStyle", "Footer"),
                        StandardTabs()),
                    Field(" PAGE  \\* MERGEFORMAT "),
                    Run(
                        RunProperties(
                            KeyValue("sz", "16")),
                        Tab(),
                        Text("Confidential Material – Copyright © Microsoft Corporation ", true)),
                    Field(" DATE  \\@ \"yyyy\"  \\* MERGEFORMAT ",
                        KeyValue("sz", "16")),
                    Run(
                        RunProperties(
                            KeyValue("sz", "16")),
                        Text(". All Rights Reserved."))),
                Footer("footer4",
                    ParaProperties(
                        KeyValue("pStyle", "Footer"),
                        StandardTabs()),
                    Run(
                        RunProperties(
                            KeyValue("sz", "16")),
                        Text("Confidential Material – Copyright © Microsoft Corporation ", true)),
                    Field(" DATE  \\@ \"yyyy\"  \\* MERGEFORMAT ",
                        KeyValue("sz", "16")),
                    Run(
                        RunProperties(
                            KeyValue("sz", "16")),
                        Text(". All Rights Reserved.")),
                    Run(
                        Tab()),
                    Field(" PAGE  \\* MERGEFORMAT ")
                    ),
                Footer("footer5",
                    ParaProperties(
                        KeyValue("pStyle", "Footer"),
                        StandardTabs()),
                    Run(
                        RunProperties(
                            KeyValue("sz", "16")),
                        Text("Confidential Material – Copyright © Microsoft Corporation ", true)),
                    Field(" DATE  \\@ \"yyyy\"  \\* MERGEFORMAT ",
                        KeyValue("sz", "16")),
                    Run(
                        RunProperties(
                            KeyValue("sz", "16")),
                        Text(". All Rights Reserved.")),
                    Run(
                        Tab()),
                    Field(" PAGE  \\* MERGEFORMAT ")
                    )
            };
        }

        static XElement Header(string name, params object[] elements)
        {
            return Part($"/word/{name}.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
                null,
                new XElement(ws + "hdr",
                    new XAttribute(XNamespace.Xmlns + "w", ws.NamespaceName),
                    Para(
                        elements)));
        }

        private static XElement HeaderBorder()
        {
            return new XElement(ws + "pBdr",
                Border("bottom", "single", "4", "1", "auto"));
        }

        private static IEnumerable<XElement> Headers()
        {
            return new List<XElement>
            {
                Header("header1",
                    ParaProperties(
                        KeyValue("pStyle", "Header"),
                        HeaderBorder()),
                    Run(
                        RunProperties(
                            new XElement(ws + "b"),
                            new XElement(ws + "bCs")),
                        Text("Table of Contents"))),
                Header("header2",
                    ParaProperties(
                        KeyValue("pStyle", "Header"),
                        HeaderBorder(),
                        KeyValue("jc", "right")),
                    Run(
                        RunProperties(
                            new XElement(ws + "b"),
                            new XElement(ws + "bCs")),
                        Text("Table of Contents"))),
                Header("header3",
                    ParaProperties(
                        KeyValue("pStyle", "Header"))),
                Header("header4",
                    ParaProperties(
                        KeyValue("pStyle", "Header"),
                        HeaderBorder()),
                    Field(" STYLEREF  \"Heading 1\" \\n  \\* MERGEFORMAT "),
                    Run(
                        RunProperties(
                            new XElement(ws + "b"),
                            new XElement(ws + "bCs")),
                        Text(".     ", true)),
                    Field(" STYLEREF  \"Heading 1\"  \\* MERGEFORMAT ")),
                Header("header5",
                    ParaProperties(
                        KeyValue("pStyle", "Header"),
                        HeaderBorder(),
                        StandardTabs(),
                        RunProperties(
                            new XElement(ws + "b"),
                            new XElement(ws + "bCs"))),
                    Run(
                        RunProperties(
                            new XElement(ws + "b"),
                            new XElement(ws + "bCs")),
                        Tab()),
                    Field(" STYLEREF  \"Heading 2\" \\n  \\* MERGEFORMAT "),
                    Run(
                        Text("     ", true)),
                    Field(" STYLEREF  \"Heading 2\"  \\* MERGEFORMAT ")),
                Header("header6",
                    ParaProperties(
                        KeyValue("pStyle", "Header"))),
            };
        }

        static XElement BulletedList(string style, string tab)
        {
            return new XElement(ws + "lvl",
                new XAttribute(ws + "ilvl", "0"),
                KeyValue("start", "1"),
                KeyValue("numFmt", "bullet"),
                KeyValue("pStyle", style),
                KeyValue("lvlText", ""),
                KeyValue("lvlJc", "left"),
                ParaProperties(
                    Tabs(
                        new Tuple<string, string>("num", tab)),
                    new XElement(ws + "ind",
                        new XAttribute(ws + "left", tab),
                        new XAttribute(ws + "hanging", "360"))),
                RunProperties(
                    new XElement(ws + "rFonts",
                        new XAttribute(ws + "ascii", "Symbol"),
                        new XAttribute(ws + "hAnsi", "Symbol"),
                        new XAttribute(ws + "hint", "default"))));
        }

        static XElement Heading(string ilvl, string style, string text, string indent)
        {
            return new XElement(ws + "lvl",
                new XAttribute(ws + "ilvl", ilvl),
                KeyValue("start", "1"),
                KeyValue("numFmt", "decimal"),
                KeyValue("pStyle", style),
                KeyValue("suff", "space"),
                KeyValue("lvlText", text),
                KeyValue("lvlJc", "left"),
                ParaProperties(
                    new XElement(ws + "ind",
                        new XAttribute(ws + "left", indent),
                        new XAttribute(ws + "hanging", indent))));
        }

        static XElement NumberedList(string style, string tab)
        {
            return new XElement(ws + "lvl",
                new XAttribute(ws + "ilvl", "0"),
                KeyValue("start", "1"),
                KeyValue("numFmt", "decimal"),
                KeyValue("pStyle", style),
                KeyValue("lvlText", "%1."),
                KeyValue("lvlJc", "left"),
                ParaProperties(
                    Tabs(
                        new Tuple<string, string>("num", tab)),
                    new XElement(ws + "ind",
                        new XAttribute(ws + "left", tab),
                        new XAttribute(ws + "hanging", "360"))),
                RunProperties(
                    new XElement(ws + "rFonts",
                        new XAttribute(ws + "hint", "default"))));
        }

        static XElement AbstractNum(string id, string nsid, params object[] lvls)
        {
            return new XElement(ws + "abstractNum",
                new XAttribute(ws + "abstractNumId", id),
                KeyValue("nsid", nsid),
                KeyValue("multiLevelType", lvls.Length > 1 ? "multilevel" : "singleLevel"),
                lvls);
        }

        static XElement Number(string id, string abstractId)
        {
            return new XElement(ws + "num",
                new XAttribute(ws + "numId", id),
                KeyValue("abstractNumId", abstractId));
        }

        static XElement Numbering()
        {
            return Part("/word/numbering.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
                null,
                new XElement(ws + "numbering",
                    new XAttribute(XNamespace.Xmlns + "w", ws.NamespaceName),
                    AbstractNum("0", "453D70D5",
                        BulletedList("BulletedList1", "360")),
                    AbstractNum("1", "70C804DC",
                        BulletedList("BulletedList2", "720")),
                    AbstractNum("2", "7AF84DA2",
                        Heading("0", "Heading1", "%1.", "432"),
                        Heading("1", "Heading2", "%1.%2", "576"),
                        Heading("2", "Heading3", "%1.%2.%3", "720"),
                        Heading("3", "Heading4", "%1.%2.%3.%4", "864"),
                        Heading("4", "Heading5", "%1.%2.%3.%4.%5", "1008"),
                        Heading("5", "Heading6", "%1.%2.%3.%4.%5.%6", "1152"),
                        Heading("6", "Heading7", "%1.%2.%3.%4.%5.%6.%7", "1296"),
                        Heading("7", "Heading8", "%1.%2.%3.%4.%5.%6.%7.%8", "1440"),
                        Heading("8", "Heading9", "%1.%2.%3.%4.%5.%6.%7.%8.%9", "1584")),
                    AbstractNum("3", "0B086C79",
                        NumberedList("NumberedList1", "360")),
                    AbstractNum("4", "49917801",
                        NumberedList("NumberedList2", "720")),
                    Number("1", "0"),
                    Number("2", "1"),
                    Number("3", "2"),
                    Number("4", "3"),
                    Number("5", "4")));
        }

        static XElement Settings()
        {
            return Part("/word/settings.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
                null,
                new XElement(ws + "settings",
                    new XAttribute(XNamespace.Xmlns + "w", ws.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "m", m.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "o", o.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "v", v.NamespaceName),
                    KeyValue("characterSpacingControl", "doNotCompress"),
                    KeyValue("decimalSymbol", "."),
                    KeyValue("defaultTabStop", "360"),
                    new XElement(ws + "evenAndOddHeaders"),
                    KeyValue("listSeparator", ","),
                    new XElement(m + "mathPr",
                        KeyValue(m, "mathFont", "Cambria Math"),
                        KeyValue(m, "brkBin", "before"),
                        KeyValue(m, "brkBinSub", "--"),
                        KeyValue(m, "smallFrac", "0"),
                        new XElement(m + "dispDef"),
                        KeyValue(m, "lMargin", "0"),
                        KeyValue(m, "rMargin", "0"),
                        KeyValue(m, "defJc", "centerGroup"),
                        KeyValue(m, "wrapIndent", "1440"),
                        KeyValue(m, "intLim", "subSup"),
                        KeyValue(m, "naryLim", "undOvr")
                    ),
                    new XElement(ws + "shapeDefaults",
                        new XElement(o + "shapedefaults",
                            new XAttribute(v + "ext", "edit"),
                            new XAttribute("spidmax", "1026")),
                        new XElement(o + "shapelayout",
                            new XAttribute(v + "ext", "edit"),
                            new XElement(o + "idmap",
                                new XAttribute(v + "ext", "edit"),
                                new XAttribute("data", "1")))),
                new XElement(ws + "zoom", 
                    new XAttribute(ws + "percent", "100"))
                ));
        }

        static XElement NumProperties(string id, string ilvl = null)
        {
            return new XElement(ws + "numPr",
                ilvl != null ? KeyValue("ilvl", ilvl) : null,
                KeyValue("numId", id));
        }

        static XElement Style(bool paragraph, bool custom, bool def, string id, string name, bool quick, string basedOn, string next, bool hidden, bool redefine, string link, params object[] props)
        {
            return new XElement(ws + "style",
                new XAttribute(ws + "type", paragraph ? "paragraph" : "character"),
                custom ? new XAttribute(ws + "customStyle", "1") : null,
                def ? new XAttribute(ws + "default", "1") : null,
                new XAttribute(ws + "styleId", id),
                KeyValue("name", name),
                quick ? new XElement(ws + "qFormat") : null,
                basedOn != null ? KeyValue("basedOn", basedOn) : null,
                next != null ? KeyValue("next", next) : null,
                hidden ? new XElement(ws + "semiHidden") : null,
                hidden ? new XElement(ws + "unhideWhenUsed") : null, 
                redefine ? new XElement(ws + "autoRedefine") : null,
                link != null ? KeyValue("link", link) : null,
                props);
        }

        static XElement Border(string type, string stroke, string size, string space, string color, string shadow = null)
        {
            return new XElement(ws + type,
                new XAttribute(ws + "val", stroke),
                new XAttribute(ws + "sz", size),
                new XAttribute(ws + "space", space),
                new XAttribute(ws + "color", color),
                shadow != null ? new XAttribute(ws + "shadow", shadow) : null);
        }

        static XElement Styles()
        {
            return Part("/word/styles.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
                null,
                new XElement(ws + "styles",
                    new XAttribute(XNamespace.Xmlns + "w", ws.NamespaceName),
                    new XElement(ws + "docDefaults",
                        new XElement(ws + "rPrDefault",
                            RunProperties(
                                new XElement(ws + "rFonts",
                                    new XAttribute(ws + "ascii", "Times New Roman"),
                                    new XAttribute(ws + "eastAsia", "Times New Roman"),
                                    new XAttribute(ws + "hAnsi", "Times New Roman"),
                                    new XAttribute(ws + "cs", "Times New Roman")),
                                new XElement(ws + "lang",
                                    new XAttribute(ws + "val", "en-US"),
                                    new XAttribute(ws + "eastAsia", "en-US"),
                                    new XAttribute(ws + "bidi", "ar-SA")))),
                        new XElement(ws + "pPrDefault")),
                    Style(true, true, false, "Annotation", "Annotation", true, "Text", null, false, false, null,
                        ParaProperties(
                            new XElement(ws + "pBdr",
                                Border("top", "single", "4", "1", "auto", "1"),
                                Border("left", "single", "4", "4", "auto", "1"),
                                Border("bottom", "single", "4", "1", "auto", "1"),
                                Border("right", "single", "4", "4", "auto", "1")),
                            new XElement(ws + "shd",
                                new XAttribute(ws + "val", "pct50"),
                                new XAttribute(ws + "color", "C0C0C0"),
                                new XAttribute(ws + "fill", "auto"))
                        )
                    ),
                    Style(true, true, false, "AlertText", "Alert Text", true, "Text", null, false, false, null,
                        ParaProperties(
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "360"))
                        )
                    ),
                    Style(true, true, false, "AlertTextinList1", "Alert Text in List 1", true, "TextinList1", null, false, false, null,
                        ParaProperties(
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "720"))
                        )
                    ),
                    Style(true, true, false, "AlertTextinList2", "Alert Text in List 2", true, "TextinList2", null, false, false, null,
                        ParaProperties(
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "1080"))
                        )
                    ),
                    Style(true, true, false, "Author", "Author", true, "Subtitle", null, false, false, null),
                    Style(false, true, false, "Bold", "Bold", true, "DefaultParagraphFont", null, false, false, null,
                        RunProperties(
                            new XElement(ws + "b")
                        )
                    ),
                    Style(false, true, false, "BoldItalic", "Bold Italic", true, "DefaultParagraphFont", null, false, false, null,
                        RunProperties(
                            new XElement(ws + "b"),
                            new XElement(ws + "i")
                        )
                    ),
                    Style(true, true, false, "BulletedList1", "Bulleted List 1", true, "Text", null, false, false, null,
                        ParaProperties(
                            NumProperties("1")
                        )
                    ),
                    Style(true, true, false, "BulletedList2", "Bulleted List 2", true, "Text", null, false, false, null,
                        ParaProperties(
                            NumProperties("2")
                        )
                    ),
                    Style(true, true, false, "Code", "Code", true, null, null, false, false, null,
                        ParaProperties(
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "after", "120")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "720"))
                        ),
                        RunProperties(
                            new XElement(ws + "rFonts",
                                    new XAttribute(ws + "ascii", "Consolas"), 
                                    new XAttribute(ws + "hAnsi", "Consolas")),
                            new XElement(ws + "noProof"),
                            KeyValue("color", "000080")
                        )
                    ),
                    Style(false, true, false, "CodeEmbedded", "Code Embedded", true, "DefaultParagraphFont", null, false, false, null,
                        RunProperties(
                            new XElement(ws + "rFonts",
                                    new XAttribute(ws + "ascii", "Consolas"), 
                                    new XAttribute(ws + "hAnsi", "Consolas")),
                            new XElement(ws + "noProof"),
                            KeyValue("color", "000080"),
                            KeyValue("position", "0"),
                            KeyValue("sz", "20"),
                            KeyValue("szCs", "20")
                        )
                    ),
                    Style(true, true, false, "CodeinList1", "Code in List 1", true, "Code", null, false, false, null,
                        ParaProperties(
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "1080"))
                        )
                    ),
                    Style(true, true, false, "CodeinList2", "Code in List 2", true, "Code", null, false, false, null,
                        ParaProperties(
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "1440"))
                        )
                    ),
                    Style(false, false, true, "DefaultParagraphFont", "Default Paragraph Font", false, null, null, true, false, null),
                    Style(true, false, false, "Footer", "Footer", true, "Text", null, false, false, null,
                        ParaProperties(
                            Tabs(
                                new Tuple<string, string>("center", "4320"),
                                new Tuple<string, string>("right", "8640"))
                        )
                    ),
                    Style(false, true, false, "Definition", "Definition", true, "DefaultParagraphFont", null, false, false, null,
                        RunProperties(
                            new XElement(ws + "i")
                        )
                    ),
                    Style(true, true, false, "DeprecatedText", "Deprecated Text", true, "Text", null, false, false, null),
                    Style(true, true, false, "DeprecatedTextinList1", "Deprecated Text in List 1", true, "TextinList1", null, false, false, null),
                    Style(true, true, false, "DeprecatedTextinList2", "Deprecated Text in List 2", true, "TextinList2", null, false, false, null),
                    Style(true, true, false, "Grammar", "Grammar", true, "Text", null, false, false, null,
                        ParaProperties(
                            new XElement(ws + "keepLines"),
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "line", "250"),
                                new XAttribute(ws + "lineRule", "exact")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "1080"),
                                new XAttribute(ws + "hanging", "360"))
                        ),
                        RunProperties(
                            new XElement(ws + "noProof")
                        )
                    ),
                    Style(false, true, false, "GrammarNon-Terminal", "Grammar Non-Terminal", true, "DefaultParagraphFont", null, false, false, null,
                        RunProperties(
                            new XElement(ws + "i"),
                            new XElement(ws + "iCs"),
                            new XElement(ws + "noProof")
                        )
                    ),
                    Style(false, true, false, "GrammarReference", "Grammar Reference", true, "DefaultParagraphFont", null, false, false, null,
                        RunProperties(
                            new XElement(ws + "i")
                        )
                    ),
                    Style(false, true, false, "GrammarTerminal", "Grammar Terminal", true, "DefaultParagraphFont", null, false, false, null,
                        RunProperties(
                            new XElement(ws + "rFonts",
                                    new XAttribute(ws + "ascii", "Consolas"),
                                    new XAttribute(ws + "hAnsi", "Consolas"),
                                    new XAttribute(ws + "cs", "Courier New")),
                            new XElement(ws + "noProof"),
                            KeyValue("color", "000080"),
                            KeyValue("sz", "20")
                        )
                    ),
                    Style(true, false, false, "Header", "Header", true, "Text", null, false, false, null,
                        ParaProperties(
                            Tabs(
                                new Tuple<string, string>("center", "4320"),
                                new Tuple<string, string>("right", "8640"))
                        )
                    ),
                    Style(true, true, false, "HeadingBase", "Heading Base", false, "Text", "Text", true, false, null,
                        ParaProperties(
                            new XElement(ws + "keepNext"),
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "before", "160"),
                                new XAttribute(ws + "after", "80"))
                        ),
                        RunProperties(
                            new XElement(ws + "rFonts",
                                    new XAttribute(ws + "ascii", "Cambria"), 
                                    new XAttribute(ws + "hAnsi", "Cambria")),
                            KeyValue("kern", "28"),
                            KeyValue("szCs", "20")
                        )
                    ),
                    Style(true, false, false, "Heading1", "Heading 1",true, "HeadingBase", "Text", false, false, null,
                        ParaProperties(
                            NumProperties("3"),
                            new XElement(ws + "pBdr",
                                Border("bottom", "double", "4", "8", "auto")),
                            Tabs(
                                new Tuple<string, string>("num", "360")),
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "after", "480")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "0"),
                                new XAttribute(ws + "firstLine", "0")),
                            KeyValue("jc", "right"),
                            KeyValue("outlineLvl", "0")
                        ),
                        RunProperties(
                            new XElement(ws + "b"),
                            KeyValue("sz", "48")
                        )
                    ),
                    Style(true, false, false, "Heading2", "Heading 2",true, "HeadingBase", "Text", false, false, null,
                        ParaProperties(
                            NumProperties("3", "1"),
                            Tabs(
                                new Tuple<string, string>("num", "360")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "0"),
                                new XAttribute(ws + "firstLine", "0")),
                            KeyValue("outlineLvl", "1")
                        ),
                        RunProperties(
                            new XElement(ws + "b"),
                            KeyValue("sz", "24")
                        )
                    ),
                    Style(true, false, false, "Heading3", "Heading 3",true, "HeadingBase", "Text", false, false, null,
                        ParaProperties(
                            NumProperties("3", "2"),
                            Tabs(
                                new Tuple<string, string>("num", "360")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "0"),
                                new XAttribute(ws + "firstLine", "0")),
                            KeyValue("outlineLvl", "2")
                        ),
                        RunProperties(
                            new XElement(ws + "b")
                        )
                    ),
                    Style(true, false, false, "Heading4", "Heading 4",true, "HeadingBase", "Text", false, false, null,
                        ParaProperties(
                            NumProperties("3", "3"),
                            Tabs(
                                new Tuple<string, string>("num", "360")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "0"),
                                new XAttribute(ws + "firstLine", "0")),
                            KeyValue("outlineLvl", "3")
                        )
                    ),
                    Style(true, false, false, "Heading5", "Heading 5",true, "HeadingBase", "Text", false, false, null,
                        ParaProperties(
                            NumProperties("3", "4"),
                            Tabs(
                                new Tuple<string, string>("num", "360")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "0"),
                                new XAttribute(ws + "firstLine", "0")),
                            KeyValue("outlineLvl", "4")
                        )
                    ),
                    Style(true, false, false, "Heading6", "Heading 6",true, "HeadingBase", "Text", false, false, null,
                        ParaProperties(
                            NumProperties("3", "5"),
                            Tabs(
                                new Tuple<string, string>("num", "360")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "0"),
                                new XAttribute(ws + "firstLine", "0")),
                            KeyValue("outlineLvl", "5")
                        )
                    ),
                    Style(true, false, false, "Heading7", "Heading 7",true, "HeadingBase", "Text", false, false, null,
                        ParaProperties(
                            NumProperties("3", "6"),
                            Tabs(
                                new Tuple<string, string>("num", "360")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "0"),
                                new XAttribute(ws + "firstLine", "0")),
                            KeyValue("outlineLvl", "6")
                        )
                    ),
                    Style(true, false, false, "Heading8", "Heading 8",true, "HeadingBase", "Text", false, false, null,
                        ParaProperties(
                            NumProperties("3", "7"),
                            Tabs(
                                new Tuple<string, string>("num", "360")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "0"),
                                new XAttribute(ws + "firstLine", "0")),
                            KeyValue("outlineLvl", "7")
                        )
                    ),
                    Style(true, false, false, "Heading9", "Heading 9",true, "HeadingBase", "Text", false, false, null,
                        ParaProperties(
                            NumProperties("3", "8"),
                            Tabs(
                                new Tuple<string, string>("num", "360")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "0"),
                                new XAttribute(ws + "firstLine", "0")),
                            KeyValue("outlineLvl", "8")
                        )
                    ),
                    Style(true, true, false, "Issue", "Issue", true, "Text", null, false, false, null,
                        ParaProperties(
                            new XElement(ws + "keepLines"),
                            new XElement(ws + "pBdr",
                                Border("top", "single", "4", "1", "D99694"),
                                Border("left", "single", "4", "4", "D99694"),
                                Border("bottom", "single", "4", "1", "D99694"),
                                Border("right", "single", "4", "4", "D99694")),
                            new XElement(ws + "shd",
                                new XAttribute(ws + "val", "clear"),
                                new XAttribute(ws + "color", "auto"),
                                new XAttribute(ws + "fill", "F2DCDB"))
                        ),
                        RunProperties(
                            new XElement(ws + "i"),
                            new XElement(ws + "noProof"),
                            KeyValue("szCs", "20")
                        )
                    ),
                    Style(false, true, false, "Italic", "Italic", true, "DefaultParagraphFont", null, false, false, null,
                        RunProperties(
                            new XElement(ws + "i")
                        )
                    ),
                    Style(true, true, false, "Label", "Label",true, "Text", "Text", false, false, null,
                        RunProperties(
                            new XElement(ws + "b")
                        )
                    ),
                    Style(false, true, false, "LabelEmbedded", "Label Embedded", true, "DefaultParagraphFont", null, false, false, null,
                        RunProperties(
                            new XElement(ws + "b"),
                            KeyValue("szCs", "20")
                        )
                    ),
                    Style(true, true, false, "LabelinList1", "Label in List 1",true, "TextinList1", "TextinList1", false, false, null,
                        RunProperties(
                            new XElement(ws + "b")
                        )
                    ),
                    Style(true, true, false, "LabelinList2", "Label in List 2",true, "TextinList2", "TextinList2", false, false, null,
                        RunProperties(
                            new XElement(ws + "b")
                        )
                    ),
                    Style(true, false, true, "Normal", "Normal", false, null, null, true, false, null,
                        ParaProperties(
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "after", "200"),
                                new XAttribute(ws + "line", "276"),
                                new XAttribute(ws + "lineRule", "auto"))
                        ),
                        RunProperties(
                            new XElement(ws + "rFonts",
                                    new XAttribute(ws + "ascii", "Calibri"), 
                                    new XAttribute(ws + "hAnsi", "Calibri")),
                            KeyValue("sz", "22"),
                            KeyValue("szCs", "22")
                        )
                    ),
                    Style(true, true, false, "NumberedList1", "Numbered List 1", true, "Text", null, false, false, null,
                        ParaProperties(
                            NumProperties("4")
                        )
                    ),
                    Style(true, true, false, "NumberedList2", "Numbered List 2", true, "Text", null, false, false, null,
                        ParaProperties(
                            NumProperties("5")
                        )
                    ),
                    Style(false, true, false, "Strikethrough", "Strikethrough", true, "DefaultParagraphFont", null, false, false, null,
                        RunProperties(
                            new XElement(ws + "strike"),
                            KeyValue("dstrike", "0")
                        )
                    ),
                    Style(false, true, false, "Subscript", "Subscript", true, "DefaultParagraphFont", null, false, false, null,
                        RunProperties(
                            KeyValue("vertAlign", "subscript")
                        )
                    ),
                    Style(true, false, false, "Subtitle", "Subtitle", true, "Text", null, false, false, null,
                        ParaProperties(
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "after", "60"))
                        ),
                        RunProperties(
                            new XElement(ws + "rFonts",
                                    new XAttribute(ws + "cs", "Arial")),
                            new XElement(ws + "i"),
                            KeyValue("sz", "36"),
                            KeyValue("szCs", "28")
                        )
                    ),
                    Style(false, true, false, "Superscript", "Superscript", true, "DefaultParagraphFont", null, false, false, null,
                        RunProperties(
                            KeyValue("vertAlign", "superscript")
                        )
                    ),
                    Style(true, true, false, "TableSpacing", "Table Spacing", false, "Text", "Text", true, false, null,
                        ParaProperties(
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "after", "0"),
                                new XAttribute(ws + "line", "120"),
                                new XAttribute(ws + "lineRule", "exact"))
                        ),
                        RunProperties(
                            KeyValue("color", "FF00FF"),
                            KeyValue("sz", "12")
                        )
                    ),
                    Style(true, true, false, "Text", "Text", true, null, null, false, false, "TextChar",
                        ParaProperties(
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "after", "120"))
                        ),
                        RunProperties(
                            new XElement(ws + "rFonts",
                                    new XAttribute(ws + "ascii", "Calibri"), 
                                    new XAttribute(ws + "hAnsi", "Calibri")),
                            KeyValue("color", "000000"),
                            KeyValue("sz", "22"),
                            KeyValue("szCs", "22")
                        )
                    ),
                    Style(false, true, false, "TextChar", "Text Char",true, "DefaultParagraphFont", null, false, false, "Text",
                        RunProperties(
                            new XElement(ws + "rFonts",
                                    new XAttribute(ws + "ascii", "Calibri"), 
                                    new XAttribute(ws + "hAnsi", "Calibri")),
                            KeyValue("color", "000000"),
                            KeyValue("sz", "22"),
                            KeyValue("szCs", "22")
                        )
                    ),
                    Style(true, true, false, "TextinList1", "Text in List 1", true, "Text", null, false, false, null,
                        ParaProperties(
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "360"))
                        )
                    ),
                    Style(true, true, false, "TextinList2", "Text in List 2", true, "Text", null, false, false, null,
                        ParaProperties(
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "720"))
                        )
                    ),
                    Style(true, false, false, "Title", "Title", true, "Text", null, false, false, null,
                        ParaProperties(
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "before", "240"),
                                new XAttribute(ws + "after", "60"))
                        ),
                        RunProperties(
                            new XElement(ws + "rFonts",
                                    new XAttribute(ws + "ascii", "Cambria"),
                                    new XAttribute(ws + "hAnsi", "Cambria"),
                                    new XAttribute(ws + "cs", "Arial")),
                            new XElement(ws + "b"),
                            new XElement(ws + "bCs"),
                            KeyValue("kern", "28"),
                            KeyValue("sz", "56"),
                            KeyValue("szCs", "32")
                        )
                    ),
                    Style(true, true, false, "TOCHeading", "TOC Heading", false, "Text", null, false, false, null,
                        ParaProperties(
                            new XElement(ws + "pBdr",
                                Border("bottom", "double", "4", "8", "auto")),
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "before", "160"),
                                new XAttribute(ws + "after", "480")),
                            KeyValue("jc", "right")
                        ),
                        RunProperties(
                            new XElement(ws + "rFonts",
                                    new XAttribute(ws + "ascii", "Cambria"), 
                                    new XAttribute(ws + "hAnsi", "Cambria")),
                            new XElement(ws + "b"),
                            new XElement(ws + "bCs"),
                            KeyValue("sz", "48"),
                            KeyValue("szCs", "20")
                        )
                    ),
                    Style(true, false, false, "TOC1", "TOC 1",true, "Text", "Text", false, true, null,
                        ParaProperties(
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "before", "120"))
                        ),
                        RunProperties(
                            new XElement(ws + "b"),
                            new XElement(ws + "bCs")
                        )
                    ),
                    Style(true, false, false, "TOC2", "TOC 2",true, "Text", "Text", false, true, null,
                        ParaProperties(
                            Tabs(
                                new Tuple<string, string, string>("right", "9926", "dot")),
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "after", "0")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "240"))
                        ),
                        RunProperties(
                            new XElement(ws + "noProof")
                        )
                    ),
                    Style(true, false, false, "TOC3", "TOC 3",true, "Text", "Text", false, true, null,
                        ParaProperties(
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "after", "0")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "480"))
                        ),
                        RunProperties(
                            new XElement(ws + "iCs")
                        )
                    ),
                    Style(true, false, false, "TOC4", "TOC 4",true, "Text", "Text", false, true, null,
                        ParaProperties(
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "after", "0")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "720"))
                        ),
                        RunProperties(
                            KeyValue("szCs", "21")
                        )
                    ),
                    Style(true, false, false, "TOC5", "TOC 5",true, "Text", "Text", false, true, null,
                        ParaProperties(
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "after", "0")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "960"))
                        ),
                        RunProperties(
                            KeyValue("szCs", "21")
                        )
                    ),
                    Style(true, false, false, "TOC6", "TOC 6",true, "Text", "Text", false, true, null,
                        ParaProperties(
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "after", "0")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "1200"))
                        ),
                        RunProperties(
                            KeyValue("szCs", "21")
                        )
                    ),
                    Style(true, false, false, "TOC7", "TOC 7",true, "Text", "Text", false, true, null,
                        ParaProperties(
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "after", "0")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "1440"))
                        ),
                        RunProperties(
                            KeyValue("szCs", "21")
                        )
                    ),
                    Style(true, false, false, "TOC8", "TOC 8",true, "Text", "Text", false, true, null,
                        ParaProperties(
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "after", "0")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "1680"))
                        ),
                        RunProperties(
                            KeyValue("szCs", "21")
                        )
                    ),
                    Style(true, false, false, "TOC9", "TOC 9",true, "Text", "Text", false, true, null,
                        ParaProperties(
                            new XElement(ws + "spacing",
                                new XAttribute(ws + "after", "0")),
                            new XElement(ws + "ind",
                                new XAttribute(ws + "left", "1920"))
                        ),
                        RunProperties(
                            KeyValue("szCs", "21")
                        ))));
        }

        static XElement WebSettings()
        {
            return Part("/word/webSettings.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
                null,
                new XElement(ws + "webSettings",
                    new XAttribute(XNamespace.Xmlns + "w", ws.NamespaceName),
                    new XElement(ws + "optimizeForBrowser"),
                    new XElement(ws + "relyOnVML"),
                    new XElement(ws + "allowPNG")));
        }

        private static void CheckSpecification(XDocument spec)
        {
            var tokenSymbols = new Dictionary<string, bool>();
            var tokens =
                (from t in spec.Descendants("token")
                where t.Attribute("ignore") == null || !bool.Parse(t.Attribute("ignore").Value)
                select t).ToList();
            var tokenNames = from t in tokens select t.Attribute("name").Value;

            foreach (var name in tokenNames)
            {
                if (tokenSymbols.ContainsKey(name))
                {
                    Console.WriteLine($"Error: Duplicate token name'{name}'.");
                }
                else
                {
                    tokenSymbols[name] = false;
                }
            }

            var tokenReferences =
                from t in tokens
                from nt in t.Descendants("nt")
                select nt.Value;
            var missingTokenReferences = new HashSet<string>();

            foreach (var tokenReference in tokenReferences)
            {
                if (!tokenSymbols.ContainsKey(tokenReference))
                {
                    missingTokenReferences.Add(tokenReference);
                }
                else
                {
                    tokenSymbols[tokenReference] = true;
                }
            }

            foreach (var missingToken in missingTokenReferences.OrderBy(v => v))
            {
                Console.WriteLine($"Error: Token reference to missing token '{missingToken}'.");
            }

            var syntaxSymbols = new Dictionary<string, bool>();
            var syntaxes =
                (from t in spec.Descendants("syntax")
                where t.Attribute("ignore") == null || !bool.Parse(t.Attribute("ignore").Value)
                select t).ToList();
            var syntaxNames = from t in syntaxes select t.Attribute("name").Value;

            foreach (var name in syntaxNames)
            {
                if (syntaxSymbols.ContainsKey(name))
                {
                    if (name != "start")
                    {
                        Console.WriteLine($"Error: Duplicate syntax name '{name}'.");
                    }
                    else
                    {
                        syntaxSymbols[name] = false;
                    }
                }

                if (tokenSymbols.ContainsKey(name) && name != "start")
                {
                    Console.WriteLine($"Error: Duplicate token/syntax name '{name}'.");
                }
            }

            var syntaxReferences =
                from t in syntaxes
                from nt in t.Descendants("nt")
                select nt.Value;
            var missingSyntaxReferences = new HashSet<string>();

            foreach (var syntaxReference in syntaxReferences)
            {
                if (!syntaxSymbols.ContainsKey(syntaxReference))
                {
                    if (!tokenSymbols.ContainsKey(syntaxReference))
                    {
                        missingSyntaxReferences.Add(syntaxReference);
                    }
                    else
                    {
                        tokenSymbols[syntaxReference] = true;
                    }
                }
                else
                {
                    syntaxSymbols[syntaxReference] = true;
                }
            }

            foreach (var missingSyntax in missingSyntaxReferences.OrderBy(v => v))
            {
                Console.WriteLine($"Error: Syntax reference to missing syntax '{missingSyntax}'.");
            }

            if (missingSyntaxReferences.Count > 0)
            {
                Console.WriteLine($"Error: Missing {missingSyntaxReferences.Count} syntax references.");
            }

            foreach (var tokenSymbol in tokenSymbols.Where(v => v.Key != "start" && !v.Value))
            {
                Console.WriteLine($"Error: Token '{tokenSymbol.Key}' is never referenced.");
            }

            foreach (var syntaxSymbol in syntaxSymbols.Where(v => v.Key != "start" && !v.Value))
            {
                Console.WriteLine($"Error: Syntax '{syntaxSymbol.Key}' is never referenced.");
            }

            var nameReferences = from r in spec.Descendants("ref") select r.Value;

            foreach (
                var nameReference in
                    nameReferences.Where(nr => !tokenSymbols.ContainsKey(nr) && !syntaxSymbols.ContainsKey(nr)))
            {
                Console.WriteLine($"Error: Missing name reference '{nameReference}.");
            }
        }

        static int Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: specgen <input path> <output path>");
                return 1;
            }

            var spec = XDocument.Load(args[0]);

            Console.WriteLine("Checking specification...");
            CheckSpecification(spec);
            Console.WriteLine("Checked specification...");

            Console.WriteLine("Writing specification...");

            var doc = new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XProcessingInstruction("mso-application", "progid=\"Word.Document\""),
                new XElement(Pkg + "package",
                    new XAttribute(XNamespace.Xmlns + "pkg", Pkg.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "ws", ws.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "rs", rs.NamespaceName),
                    PackageRelationships(),
                    DocumentRelationships(),
                    FontTable(),
                    Document(spec),
                    Footers(),
                    Headers(),
                    Numbering(),
                    Settings(),
                    Styles(),
                    WebSettings()));

            doc.Save(args[1]);

            return 0;
        }
    }
}
