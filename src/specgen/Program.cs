using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace specgen
{
    internal class Program
    {
        private static readonly XNamespace Pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        private static readonly XNamespace Ors = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace Prs = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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
            return new XElement(Prs + "Relationships",
                from r in relationships
                select new XElement(Prs + "Relationship",
                    new XAttribute("Id", r.Id),
                    new XAttribute("Type", $"{Ors.NamespaceName}/{r.Type}"),
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

        private static XElement KeyValue(string key, string value)
        {
            return new XElement(W + key,
                new XAttribute(W + "val", value));
        }

        private static XElement KeyValue(string key, string valueName, string value)
        {
            return new XElement(W + key,
                new XAttribute(W + valueName, value));
        }

        private static XElement Fonts(IEnumerable<Font> fonts)
        {
            return new XElement(W + "fonts",
                from f in fonts
                select new XElement(W + "font",
                    new XAttribute(W + "name", f.Name),
                    KeyValue("panose1", f.Panose1),
                    KeyValue("charset", f.Charset),
                    KeyValue("family", f.Family),
                    KeyValue("pitch", f.Pitch),
                    new XElement(W + "sig",
                        new XAttribute(W + "usb0", f.Sig.Usb0),
                        new XAttribute(W + "usb1", f.Sig.Usb1),
                        new XAttribute(W + "usb2", f.Sig.Usb2),
                        new XAttribute(W + "usb3", f.Sig.Usb3),
                        new XAttribute(W + "csb0", f.Sig.Csb0),
                        new XAttribute(W + "csb1", f.Sig.Csb1)
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

        static XElement Document()
        {
            return Part("/word/document.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                null,
                new XElement(W + "document",
                    new XElement(W + "body",
                        // TODO
                        null)));
        }

        static XElement Para(params object[] elements)
        {
            return new XElement(W + "p",
                elements);
        }

        static XElement Run(params object[] elements)
        {
            return new XElement(W + "r",
                elements);
        }

        static XElement Text(string text, bool preserve = false)
        {
            return new XElement(W + "t",
                preserve ? new XAttribute(XNamespace.Xml + "space", "preserve") : null,
                text);
        }

        static XElement Break()
        {
            return new XElement(W + "br");
        }

        static XElement Tab()
        {
            return new XElement(W + "tab");
        }

        static XElement ParaProperties(params object[] elements)
        {
            return new XElement(W + "pPr",
                elements);
        }

        static XElement RunProperties(params object[] elements)
        {
            return new XElement(W + "rPr",
                elements);
        }

        static IEnumerable<XElement> Field(string value, params object[] properties)
        {
            return new List<XElement>
            {
                Run(
                    properties != null ? RunProperties(properties) : null,
                    new XElement(W + "fldChar",
                        new XAttribute(W + "fldCharType", "begin"))),
                Run(
                    properties != null ? RunProperties(properties) : null,
                    new XElement(W + "instrText",
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        value)),
                Run(
                    properties != null ? RunProperties(properties) : null,
                    new XElement(W + "fldChar",
                        new XAttribute(W + "fldCharType", "separate"))),
                Run(
                    properties != null ? RunProperties(properties) : null,
                    new XElement(W + "fldChar",
                        new XAttribute(W + "fldCharType", "end")))
            };
        }

        static XElement Tabs(params Tuple<string, string>[] tabs)
        {
            return new XElement(W + "tabs",
                from t in tabs
                select new XElement(W + "tab",
                    new XAttribute(W + "val", t.Item1),
                    new XAttribute(W + "pos", t.Item2)));
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
                new XElement(W + "ftr",
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
                new XElement(W + "hdr",
                    Para(
                        elements)));
        }

        private static XElement HeaderBorder()
        {
            return new XElement(W + "bdr",
                new XElement(W + "bottom",
                    new XAttribute(W + "val", "single"),
                    new XAttribute(W + "sz", "4"),
                    new XAttribute(W + "space", "1"),
                    new XAttribute(W + "color", "auto")));
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
                            new XElement(W + "b"),
                            new XElement(W + "bCs")),
                        Text("Table of Contents"))),
                Header("header2",
                    ParaProperties(
                        KeyValue("pStyle", "Header"),
                        HeaderBorder(),
                        KeyValue("jc", "right")),
                    Run(
                        RunProperties(
                            new XElement(W + "b"),
                            new XElement(W + "bCs")),
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
                            new XElement(W + "b"),
                            new XElement(W + "bCs")),
                        Text(".     ", true)),
                    Field(" STYLEREF  \"Heading 1\"  \\* MERGEFORMAT ")),
                Header("header5",
                    ParaProperties(
                        KeyValue("pStyle", "Header"),
                        HeaderBorder(),
                        StandardTabs(),
                        RunProperties(
                            new XElement(W + "b"),
                            new XElement(W + "bCs"))),
                    Run(
                        RunProperties(
                            new XElement(W + "b"),
                            new XElement(W + "bCs")),
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
            return new XElement(W + "lvl",
                new XAttribute(W + "ilvl", "0"),
                KeyValue("start", "1"),
                KeyValue("numFmt", "bullet"),
                KeyValue("pStyle", style),
                KeyValue("lvlText", ""),
                KeyValue("lvlJc", "left"),
                ParaProperties(
                    Tabs(
                        new Tuple<string, string>("num", tab)),
                    new XElement(W + "ind",
                        new XAttribute(W + "left", tab),
                        new XAttribute(W + "hanging", "360"))),
                RunProperties(
                    new XElement(W + "rFonts",
                        new XAttribute(W + "ascii", "Symbol"),
                        new XAttribute(W + "hAnsi", "Symbol"),
                        new XAttribute(W + "hint", "default"))));
        }

        static XElement Heading(string ilvl, string style, string text, string indent)
        {
            return new XElement(W + "lvl",
                new XAttribute(W + "ilvl", ilvl),
                KeyValue("start", "1"),
                KeyValue("numFmt", "decimal"),
                KeyValue("pStyle", style),
                KeyValue("suff", "space"),
                KeyValue("lvlText", text),
                KeyValue("lvlJc", "left"),
                RunProperties(
                    new XElement(W + "ind",
                        new XAttribute(W + "left", indent),
                        new XAttribute(W + "hanging", indent))));
        }

        static XElement NumberedList(string style, string tab)
        {
            return new XElement(W + "lvl",
                new XAttribute(W + "ilvl", "0"),
                KeyValue("start", "1"),
                KeyValue("numFmt", "decimal"),
                KeyValue("pStyle", style),
                KeyValue("lvlText", "%1."),
                KeyValue("lvlJc", "left"),
                ParaProperties(
                    Tabs(
                        new Tuple<string, string>("num", tab)),
                    new XElement(W + "ind",
                        new XAttribute(W + "left", tab),
                        new XAttribute(W + "hanging", "360"))),
                RunProperties(
                    new XElement(W + "rFonts",
                        new XAttribute(W + "hint", "default"))));
        }

        static XElement AbstractNum(string id, string nsid, params object[] lvls)
        {
            return new XElement(W + "abstractNum",
                new XAttribute(W + "abstractNumId", id),
                KeyValue("nsid", nsid),
                KeyValue("multiLevelType", lvls.Length > 1 ? "multilevel" : "singleLevel"),
                lvls);
        }

        static XElement Number(string id, string abstractId)
        {
            return new XElement(W + "num",
                new XAttribute(W + "numId", id),
                KeyValue("abstractNumId", abstractId));
        }

        static XElement Numbering()
        {
            return Part("/word/numbering.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
                null,
                new XElement(W + "numbering",
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

        static int Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: specgen <input path> <output path>");
                return 1;
            }

            var doc = new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XProcessingInstruction("mso-application", "progid=\"Word.Document\""),
                new XElement(Pkg + "package",
                    new XAttribute(XNamespace.Xmlns + "pkg", Pkg.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "ors", Ors.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "prs", Prs.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName),
                    PackageRelationships(),
                    DocumentRelationships(),
                    FontTable(),
                    Document(),
                    Footers(),
                    Headers(),
                    Numbering()));

            doc.Save(args[1]);

            return 0;
        }
    }
}
