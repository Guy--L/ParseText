using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ParseText
{
    class Reading
    {
        public double stress;
        public double strain;
        public double prime;
        public double dprime;

        // overload reading members to provide two facades
        public double time
        {
            get { return stress; }
            set { stress = value; }
        }
        public double rate
        {
            get { return strain; }
            set { stress = value; }
        }
        public double normal
        {
            get { return prime; }
            set { prime = value; }
        }
        public double shear
        {
            get { return dprime; }
            set { dprime = value; }
        }
        public Reading(string val)
        {
            if (string.IsNullOrWhiteSpace(val))
                return;

            var a = val.Split('\t');
            stress = double.Parse(a[0]);
            strain = double.Parse(a[1]);
            prime = double.Parse(a[2]);
            dprime = double.Parse(a[3]);
        }
        public Reading(double t, double n)
        {
            time = t;
            normal = n;
        }
        public Reading(double t, double n, double b)
        {
            time = t;
            normal = Math.Abs(n) > b ? 0 : n;
        }
        public Reading(Reading toZero)
        {
            rate = 0.0;
            shear = 0.0;
            time = toZero.time;
            normal = toZero.normal;
        }
        public Reading(Reading a, Reading b)
        {
            time = a.time;
            rate = b.time;

            normal = (b.normal - a.normal) / (b.time - a.time);
        }
        public Reading cutoff(double threshold)
        {
            if (normal > threshold) normal = 0.0;
            return this;
        }
        public string print()
        {
            return "(" + stress + ", " + strain + ", " + prime + ", " + dprime + ")";
        }
    }
}
