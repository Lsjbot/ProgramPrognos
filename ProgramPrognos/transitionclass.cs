using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProgramPrognos
{
    public class transitionclass
    {
        public double transitionprob;
        public double transitionsig;
        private static Random rnd = new Random();

        public transitionclass(double tprob,double tsig)
        {
            this.transitionprob = tprob;
            this.transitionsig = tsig;
        }

        public double nextnum(double oldnum,bool spread, bool poisson)
        {
            if (spread)
            {
                if (poisson)
                {
                    int newnum = 0;
                    for (int i=1;i<(oldnum+0.5);i++)
                    {
                        if (rnd.NextDouble() < transitionprob)
                            newnum++;
                    }
                    return newnum;
                }
                else
                    return oldnum * util.SampleGaussian(rnd, transitionprob, transitionsig);
            }
            else
                return oldnum * transitionprob;
        }

        public static transitionclass[] average(List<transitionclass[]> qt) //creates a new transitionclass that is average of the list
        {
            transitionclass[] tc = new transitionclass[qt.First().Length];
            for (int i=0;i<tc.Length;i++)
            {
                if (qt.First()[i] == null)
                    break;
                tc[i] = new transitionclass(0, 0);
                foreach (transitionclass[] tc2 in qt)
                {
                    tc[i].transitionprob += tc2[i].transitionprob;
                    tc[i].transitionsig += tc2[i].transitionsig;
                }
                tc[i].transitionprob /= qt.Count;
                tc[i].transitionsig /= qt.Count;
            }

            return tc;
        }
    }
}
