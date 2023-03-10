using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProgramPrognos
{
    public class institutionclass
    {
        public string name = "";
        public string shortname = "";

        public fracprodclass baseyearprod = new fracprodclass(); //base year production
        public Dictionary<int, fracprodclass> yearproddict = new Dictionary<int, fracprodclass>();
        public Dictionary<int, forecastrangeclass> yearprodrangedict = new Dictionary<int, forecastrangeclass>();
        public Dictionary<int, Dictionary<string, fracprodclass>> progyearproddict = new Dictionary<int, Dictionary<string, fracprodclass>>();
        public Dictionary<int, Dictionary<string, forecastrangeclass>> progyearprodrangedict = new Dictionary<int, Dictionary<string, forecastrangeclass>>();

        public transitionclass[] meantransition = new transitionclass[programbatchclass.maxsem];   //transition from Tn to Tn+1, average for programs
        public transitionclass[] meanexamtransition = new transitionclass[programbatchclass.maxsem]; //transition from Tn to exam, average for programs


        //public Dictionary<int, double> hstdict = new Dictionary<int, double>();
        //public Dictionary<int, double> hprdict = new Dictionary<int, double>();
        //public Dictionary<int, double> moneydict = new Dictionary<int, double>();
        //public Dictionary<int, Dictionary<string,double>> hstprogdict = new Dictionary<int, Dictionary<string, double>>();
        //public Dictionary<int, Dictionary<string, double>> hprprogdict = new Dictionary<int, Dictionary<string, double>>();
        //public Dictionary<int, Dictionary<string, double>> moneyprogdict = new Dictionary<int, Dictionary<string, double>>();

        public institutionclass(string namepar)
        {
            this.name = namepar;
            this.shortname = Form1.instshortdict[this.name];
        }

        public void calculate_meantransition()
        {
            double[] sumtrans = new double[programbatchclass.maxsem];
            double[] sumexamtrans = new double[programbatchclass.maxsem];
            int[] ntrans = new int[programbatchclass.maxsem];
            int[] nexamtrans = new int[programbatchclass.maxsem];

            for (int i=0;i<programbatchclass.maxsem;i++)
            {
                sumtrans[i] = 0;
                sumexamtrans[i] = 0;
                ntrans[i] = 0;
                nexamtrans[i] = 0;
            }

            foreach (programclass pc in Form1.origprogramdict.Values)
            {
                if (pc.homeinst != this.name)
                    continue;
                for (int i=0;i<programbatchclass.maxsem;i++)
                {
                    if (pc.transition[i] != null)
                    {
                        sumtrans[i] += pc.transition[i].transitionprob;
                        ntrans[i]++;
                    }
                    if (pc.examtransition[i] != null)
                    {
                        sumexamtrans[i] += pc.examtransition[i].transitionprob;
                        nexamtrans[i]++;
                    }
                }
            }

            for (int i = 0; i < programbatchclass.maxsem; i++)
            {
                if (ntrans[i] > 0)
                {
                    meantransition[i] = new transitionclass(sumtrans[i] / ntrans[i],0);
                }
                if (nexamtrans[i] > 0)
                {
                    meanexamtransition[i] = new transitionclass(sumexamtrans[i] / nexamtrans[i], 0);
                }
            }

        }

        public void clearproduction(int baseyear, int endyear)
        {
            for (int year=baseyear;year <= endyear; year++)
            {
                if (yearproddict.ContainsKey(year))
                    yearproddict[year] = new fracprodclass();
                else
                    yearproddict.Add(year, new fracprodclass());
                if (progyearproddict.ContainsKey(year))
                    progyearproddict[year] = new Dictionary<string, fracprodclass>();
                else
                    progyearproddict.Add(year, new Dictionary<string, fracprodclass>());
                if (yearprodrangedict.ContainsKey(year))
                    yearprodrangedict[year] = new forecastrangeclass();
                else
                    yearprodrangedict.Add(year, new forecastrangeclass());
                if (progyearprodrangedict.ContainsKey(year))
                    progyearprodrangedict[year] = new Dictionary<string, forecastrangeclass>();
                else
                    progyearprodrangedict.Add(year, new Dictionary<string, forecastrangeclass>());
            }
        }

        public void addproduction(int year, string prog, fracprodclass prod)
        {
            yearproddict[year].add(prod);
            //yearprodrangedict[year].Add(prod);
            //yearprodrangedict[year].Add(prod.fracmoney);
            if (!progyearproddict[year].ContainsKey(prog))
            {
                progyearproddict[year].Add(prog, prod.clone());
                progyearprodrangedict[year].Add(prog, new forecastrangeclass());
            }
            else
            {
                progyearproddict[year][prog].add(prod);
            }
            progyearprodrangedict[year][prog].Add(prod);
            progyearprodrangedict[year][prog].Add(prod.fracmoney);
        }

        public void scaleproduction(int year, double downscalefactor)
        {
            yearproddict[year].normalize(downscalefactor);
            foreach (string prog in progyearproddict[year].Keys)
                progyearproddict[year][prog].normalize(downscalefactor);
        }

        public void scaleproduction(double downscalefactor) //divides by downscalefactor
        {
            foreach (int year in yearproddict.Keys)
            {
                scaleproduction(year,downscalefactor);
            }
        }

        public void addproductionrange(int year,string prog, forecastrangeclass prodrange)
        {
            foreach (fracprodclass f in prodrange.fpc)
                addproduction(year, prog, f);
            yearprodrangedict[year].AddRange(prodrange);
        }

        public void addbaseyearproduction(double hst,double hpr,double hstkr, double hprkr, double kr)
        {
            baseyearprod.add(hst, hpr, hstkr, hprkr, kr);
            //baseyearprod.frachst += hst;
            //baseyearprod.frachpr += hpr;
            //baseyearprod.frachstmoney += hstkr;
            //baseyearprod.frachprmoney += hprkr;
            //baseyearprod.fracmoney += kr;

        }
    }

}
