using System;
using System.Collections.Generic;
using System.Linq;
using mn = MathNet.Numerics;
using System.Text;

namespace ParseText
{
    class Line
    {
        private double[] x;
        private double[] y;

        private IEnumerable<Reading> _data;
        public double intercept;
        public double slope;

        public Line(IEnumerable<Reading> r) : this(r, c => c.stress)
        { }

        public Line(IEnumerable<Reading> r, Func<Reading, double> column)
        {
            _data = r;
            //_data.ForEach(d => Console.WriteLine(column(d) + ", " + d.strain));
            Fit(column);
            //Console.WriteLine("y=" + slope + "x+" + intercept);
        }

        public double Fit(Func<Reading, double> column)
        {
            y = _data.Select(s => column(s)).ToArray();
            x = _data.Select(s => s.strain).ToArray();
            Tuple<double, double> z = mn.Fit.Line(x, y);

            intercept = z.Item1;
            slope = z.Item2;
            return 0.0;
        }
    }
}
