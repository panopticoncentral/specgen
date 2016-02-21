using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace specgen
{
    internal class Program
    {
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

            foreach (var missingToken in missingTokenReferences.OrderBy(value => value))
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

            foreach (var missingSyntax in missingSyntaxReferences.OrderBy(value => value))
            {
                Console.WriteLine($"Error: Syntax reference to missing syntax '{missingSyntax}'.");
            }

            if (missingSyntaxReferences.Count > 0)
            {
                Console.WriteLine($"Error: Missing {missingSyntaxReferences.Count} syntax references.");
            }

            foreach (var tokenSymbol in tokenSymbols.Where(value => value.Key != "start" && !value.Value))
            {
                Console.WriteLine($"Error: Token '{tokenSymbol.Key}' is never referenced.");
            }

            foreach (var syntaxSymbol in syntaxSymbols.Where(value => value.Key != "start" && !value.Value))
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

        private static int Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: specgen <input path> <XML output path>");
                return 1;
            }

            var spec = XDocument.Load(args[0]);

            Console.WriteLine("Checking specification...");
            CheckSpecification(spec);
            Console.WriteLine("Checked specification...");

            Console.WriteLine("Writing specification...");

            XmlSpecificationWriter.WriteSpecification(spec, args[1]);

            return 0;
        }
    }
}
