using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NewParse
{
    class Program
    {
        static void Main(string[] args)
        {
            var files = Directory.GetFiles(args[0], "*.txt");
            foreach (var file in files)
            {
                Console.WriteLine(Path.GetFileNameWithoutExtension(file));
                var lines = File.ReadAllLines(file);
                int last = 0;
                var count = lines.Select((s, i) => new { line = s, num = i }).Where(s => s.line == "[step]").Select(s => { Console.WriteLine(s.num-last); last = s.num; return s.num; }).ToList();
                Console.WriteLine(lines.Count()-last);
            }
            Console.ReadKey();
        }
    }
}
