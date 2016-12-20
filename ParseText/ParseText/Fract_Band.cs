using System.Linq;

namespace ParseText
{
    class Fract_Band : Test
    {
        public Fract_Band()
        {
            targetrows = 418;
        }

        public override void Analyze()
        {
            var orig = lines.Skip(firstline).Take(targetrows).Select(s => new Reading(s)).ToList();
            var data = orig.Select(d => (d.rate >= 0.95 && d.rate <= 1.05) ? d : new Reading(d));

            var max = data.Take(319).Max(s => s.shear);
            var sam = data.Take(319).TakeWhile(s => s.shear < max);

            var start = sam.Any() ? sam.Count() : 319;
            var avg = data.Skip(start).Take(101).Average(s => s.shear);
            var category = avg < max;

            var at5 = data.Last(t => t.time <= 5.0).shear;

            var span = data.Max(t => t.time);
            var at60 = span >= 60.0 ? data.First(t => t.time >= 60.0).shear : data.Last().shear;
            var delta = at60 - at5;

            outrow.Cell(23).SetValue<int>(category ? 1 : 0);
            outrow.Cell(24).SetValue<double>(at5);
            outrow.Cell(25).SetValue<double>(at60);
            outrow.Cell(26).SetValue<double>(delta);
            if (category) outrow.Cell(27).SetValue<double>(orig.Max(s => s.shear));
        }
    }
}
