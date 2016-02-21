using System.IO;
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
            writer.WriteLine("<br/>");
        }

        public static void WriteSpecification(XDocument spec, string path)
        {
            using (var stream = new FileStream(path, FileMode.Create))
            {
                using (var writer = new StreamWriter(stream))
                {
                    TitleSection(spec, writer);
                }
            }
        }
    }
}
