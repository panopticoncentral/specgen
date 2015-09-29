using System;
using System.Xml.Linq;

namespace specgen
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: specgen <input path> <output path>");
                return 1;
            }

            var pkg = (XNamespace)"http://schemas.microsoft.com/office/2006/xmlPackage";

            var doc = new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XProcessingInstruction("mso-application", "progid=\"Word.Document\""),
                new XElement(pkg + "package",
                    new XAttribute(XNamespace.Xmlns + "pkg", pkg.NamespaceName)));

            doc.Save(args[1]);

            return 0;
        }
    }
}
