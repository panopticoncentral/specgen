using System;

namespace specgen
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length > 0)
            {
                Console.WriteLine("specgen: Invalid command-line arguments.");
                return 1;
            }

            return 0;
        }
    }
}
