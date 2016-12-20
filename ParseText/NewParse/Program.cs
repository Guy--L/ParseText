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
        private static string sub(string line)
        {
            string s = System.Text.RegularExpressions.Regex.Replace(line, @"\s+", " ");
            return s.Substring(0, s.Length > 39 ? 39 : s.Length);
        }

        private static void check(IEnumerable<string> chunk)
        {
            var pos = 0;
            Console.WriteLine($"{pos,-3}{sub(chunk.First())}");
            var tail = chunk.Last(); 
                
                //chunk.Skip(chunk.Count() - 3).Take(1).Single();
            Console.WriteLine($"{chunk.Count(),19} {sub(tail)}");
        }

        private static int take(int type, int cnt)
        {
            var std = Test.rowmap[type];
            return std > cnt ? cnt : std;
        }

        static void Main(string[] args)
        {
            var files = Directory.GetFiles(args[0], "*.txt");
            foreach (var file in files)
            {
                Console.WriteLine(Path.GetFileNameWithoutExtension(file));
                var lines = File.ReadAllLines(file);
                int last = 0;
                var count = lines.Select((s, i) => new { line = s, num = i })
                    .Where(s => s.line == "[step]")
                    .Select(s => { var cnt = s.num-last-4; last = s.num; return cnt; }).ToList();
                count.Add(lines.Count() - last - 4);
                var pos = 4;
                foreach(var type in count)
                {
                    var typed = Test.typer(type);
                    Console.Write($"{type,-4} {typed.ToString(),-12}");
                    if (typed != Test.TestType.Trim && typed != Test.TestType.All)
                        check(lines.Skip(pos).Take(take((int)typed, type-1)));
                    else
                        Console.WriteLine();
                    pos += type+4;
                }
                Console.WriteLine();
            }
            Console.ReadKey();
        }
    }
}
