using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProgramPrognos
{
    public class forecastrangeclass
    {
        private List<double> fc = new List<double>();
        public List<fracprodclass> fpc = new List<fracprodclass>();
        private double sum = 0;
        private double sum2 = 0;
        private double xmax = double.NegativeInfinity;
        private double xmin = double.PositiveInfinity;

        public void Add(double x)
        {
            fc.Add(x);
            sum += x;
            sum2 += x * x;
            if (x > xmax)
                xmax = x;
            if (x < xmin)
                xmin = x;
        }

        public void Add(fracprodclass xf)
        {
            fpc.Add(xf);
        }

        public void AddRange(forecastrangeclass frc)
        {
            if (fc.Count() < frc.fc.Count())
            {
                for (int i = fc.Count(); i < frc.fc.Count(); i++)
                    fc.Add(0);
            }
            if (fpc.Count() < frc.fpc.Count())
            {
                for (int i = fpc.Count(); i < frc.fpc.Count(); i++)
                    fpc.Add(new fracprodclass());
            }

            for (int i=0;i<fc.Count;i++)
            {
                fc[i] += frc.fc[i];
                fpc[i].add(frc.fpc[i]);
            }
        }

        public double Average()
        {
            return sum / fc.Count;
        }

        public double Sigma()
        {
            double sig = Math.Sqrt(sum2 * fc.Count - sum * sum) / fc.Count;
            return sig;
        }

        public void SetMinMax()
        {
            foreach (double x in fc)
            {
                if (x > xmax)
                    xmax = x;
                if (x < xmin)
                    xmin = x;
            }
        }

        public Tuple<double,double> Range()
        {
            SetMinMax();
            return new Tuple<double, double>( xmin, xmax );
        }

        public string RangeString()
        {
            SetMinMax();
            return xmin + " - " + xmax;
        }
    }
}
