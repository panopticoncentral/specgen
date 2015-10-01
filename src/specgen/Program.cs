using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace specgen
{
    class Program
    {
        private static readonly XNamespace Pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        private static readonly XNamespace Ors = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace Prs = "http://schemas.openxmlformats.org/package/2006/relationships";

        private struct Relationship
        {
            public string Id;
            public string Type;
            public string Target;
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
                    PackageRelationships(),
                    DocumentRelationships()));

            doc.Save(args[1]);

            return 0;
        }
    }
}
