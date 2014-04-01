using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ezip
{
    class Program
    {
        static void Main(string[] args)
        {
            var files = new List<string>();
            var target = "";
            var overwrite = false;
            var scale = 1.0;
            for (var i = 0; i < args.Length; ++i)
            {
                switch (args[i])
                {
                    case "-o":
                        if (args.Length == (i + 1)) goto Error;
                        target = args[++i];
                        break;
                    case "-O":
                        overwrite = true;
                        break;
                    case "-s":
                        if (args.Length == (i + 1)) goto Error;
                        if (!Double.TryParse(args[++i], out scale)) goto Error;
                        break;
                    default:
                        files.Add(args[i]);
                        break;
                }
            }
            if (string.IsNullOrWhiteSpace(target)) goto Error;
            if (files.Count == 0) goto Error;
            if (File.Exists(target))
            {
                if (overwrite)
                {
                    File.Delete(target);
                }
                else
                {
                    Console.WriteLine("File '{0}' is already exists.", target);
                    return;
                }
            }
            var engine = new Engine { ImageScaling = scale };
            engine.Compress(target, files);
            return;

        Error:
            Console.WriteLine("usage: ezip -o target file [file [file...]]");
            Console.WriteLine("       -o  Output file name.");
            Console.WriteLine("       -O  Overwrite output file.");
            Console.WriteLine("       -s  Image scale");
        }
    }
}
