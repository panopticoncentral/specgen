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

        private static XElement PackageRelationships()
        {
            return new XElement(Pkg + "part",
                new XAttribute(Pkg + "name", "/_rels/.rels"),
                new XAttribute(Pkg + "contentType", "application/vnd.openxmlformats-package.relationships+xml"),
                new XAttribute(Pkg + "padding", "512"),
                new XElement(Pkg + "xmlData",
                    CreateRelationships(new List<Relationship>
                    {
                        new Relationship { Id = "rId1", Type = "officeDocument", Target = "word/document.xml" }
                    })));
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
                    PackageRelationships()));

            doc.Save(args[1]);

            return 0;
        }
    }
}
