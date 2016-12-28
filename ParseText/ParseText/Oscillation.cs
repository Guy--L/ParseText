using System.Diagnostics;
using System.Linq;

namespace ParseText
{
    class Oscillation : Test
    {
        private static int initialskip = 4;
        private static int initialtake = 5;
        private static int finalskip = 32;
        private static int finaltake = 3;

        public Oscillation()
        {
            targetrows = 38;
        }

        public override void Analyze()
        {
            var readings = lines.Skip(firstline).Select(s => new Reading(s));
            Debug.WriteLine(lines.Skip(firstline).First());
            Debug.WriteLine("readings " + readings.Count());
            var initial = readings.Skip(initialskip).Take(initialtake);
            var final = readings.Skip(finalskip).Take(finaltake);

            var iline = new Line(initial);
            var fline = new Line(final);

            var intersectx = (iline.intercept - fline.intercept) / (fline.slope - iline.slope);
            //Console.WriteLine("intersectx: " + intersectx);

            var mid = readings.TakeWhile(s => s.strain < intersectx);
            var midtriple = readings.Skip(mid.Count() - 2).Take(3);
            var mline = new Line(midtriple);

            var ypstrain = (iline.intercept - mline.intercept) / (mline.slope - iline.slope);
            var ypt = readings.TakeWhile(s => s.strain < ypstrain);
            Debug.WriteLine("ypt " + ypt.Count());
            var ypi = readings.Skip(ypt.Count() - 1).Take(2);
            if (ypi.Count() < 2)
            {
                ypi = readings.Skip(ypt.Count() - 2).Take(2);
                Debug.WriteLine("redid to " + (ypt.Count() - 2));
            }
            var yp = ypi.ToArray();

            var ypstress = (yp[1].stress - yp[0].stress) / (yp[1].strain - yp[0].strain) * (ypstrain - yp[0].strain) + yp[0].stress;
            var bpstrain = (fline.intercept - mline.intercept) / (mline.slope - fline.slope);
            var bpt = readings.TakeWhile(s => s.strain < bpstrain);
            var bp = readings.Skip(bpt.Count() - 1).Take(2).ToArray();
            var bpstress = (bp[1].stress - bp[0].stress) / (bp[1].strain - bp[0].strain) * (bpstrain - bp[0].strain) + bp[0].stress;

            var cross = readings.Skip(1).TakeWhile(s => s.prime >= s.dprime);
            Debug.WriteLine("cross " + cross.Count());
            var crx = readings.Skip(cross.Count()).Take(2);
            var cr = crx.ToArray();

            var pline = new Line(cr, c => c.prime);
            var dline = new Line(cr, c => c.dprime);

            var strd = cr[1].strain - cr[0].strain;
            var denom = ((cr[1].dprime - cr[0].dprime) / strd - (cr[1].prime - cr[0].prime) / strd);

            var gstrain = (pline.intercept - dline.intercept) / denom;
            var gstress = (cr[1].prime - cr[0].prime) / strd * (gstrain - cr[1].strain) + cr[1].prime;

            var gplateau = initial.Average(g => g.prime);
            var dplateau = initial.Average(g => g.dprime);

            outrow.Cell(15).SetValue(ypstress);
            outrow.Cell(16).SetValue(ypstrain);
            outrow.Cell(17).SetValue(bpstress);
            outrow.Cell(18).SetValue(bpstrain);

            outrow.Cell(19).SetValue(gstress);
            outrow.Cell(20).SetValue(gstrain);
            outrow.Cell(21).SetValue(gplateau);
            outrow.Cell(22).SetValue(dplateau);
        }
    }
}
