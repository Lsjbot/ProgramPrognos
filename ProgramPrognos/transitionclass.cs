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
            for (int i = 0; i < tc.Length; i++)
            {
                if (qt.First()[i] == null)
                    break;
                tc[i] = new transitionclass(0, 0);
                int nt = 0;
                foreach (transitionclass[] tc2 in qt)
                {
                    if (tc2[i] != null)
                    {
                        tc[i].transitionprob += tc2[i].transitionprob;
                        tc[i].transitionsig += tc2[i].transitionsig;
                        nt++;
                    }
                }
                if (nt > 0)
                {
                    tc[i].transitionprob /= nt;
                    tc[i].transitionsig /= nt;
                }
                else
                    tc[i].transitionprob = 0.8; ;

            }

            return tc;
        }

        public transitionclass clone()
        {
            return new transitionclass(this.transitionprob, this.transitionsig);
        }

        public static transitionclass[] clone(transitionclass[] oldarray)
        {
            if (oldarray == null)
                return null;
            transitionclass[] newarray = new transitionclass[oldarray.Length];
            for (int i = 0; i < oldarray.Length; i++)
            {
                if (oldarray[i] == null)
                    newarray[i] = null;
                else
                    newarray[i] = oldarray[i].clone();
            }
            return newarray;
        }

        public static Dictionary<int, transitionclass[]> clone(Dictionary<int, transitionclass[]> olddict)
        {
            Dictionary<int, transitionclass[]> newdict = new Dictionary<int, transitionclass[]>();
            foreach (int key in olddict.Keys)
                newdict.Add(key, transitionclass.clone(olddict[key]));
            return newdict;
        }
    }
}
