using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProgramPrognos
{
    public class programbatchclass
    {
        public static int maxsem = 20;
        public static int maxid = 0;
        public static int examforecastsem = 1;

        public string batchstart = "HT21";
        public int id = -1;
        public int progid = -1;
        public double?[] actualsemstud = new double?[] { null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }; //semstud[0] = accepted
        public double?[] forecastsemstud = new double?[] { null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null }; //semstud[0] = accepted
        public static double?[] nulldouble = new double?[] { null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null };
        public double? actualexam = null;
        public double? forecastexam = null;
        //0=appl, 1=U1, 2=U2, 3=final accepted
        public double?[] applicants = new double?[4] { null,null,null,null };  
        public double reserves = 0;
        public double budget_T1 = 0;
        public bool actualbatch = true; //true for batch with real data, false for pure forecast
        internal dictclass appldict = null;
        public double progstud = 0; //students in course that come from programs.
        public double exchangestud = 0; //students in course from Erasmus etc. Not included in actualsemstud

        public programbatchclass cloneactual() //only clone batches with real data, not forecast 
        {
            if (!actualbatch)
                return null;

            return new programbatchclass(this.actualsemstud,this.progid,this.batchstart,(int)this.actualexam,this.applicants, (int)this.reserves);
            
        }

        public programbatchclass(double accepted, int prog, string bstart, transitionclass[] transition, transitionclass[] examtransition) //forecast
        {
            maxid++;
            this.id = maxid;
            this.progid = prog;
            this.actualbatch = false;
            this.batchstart = util.semester4to2(bstart);
            //Array.Copy(nulldouble,actualsemstud,maxsem);
            //Array.Copy(nulldouble, actualsemstud, maxsem);
            forecastsemstud[0] = accepted;
            int i = 0;
            while (transition[i] != null)
            {
                forecastsemstud[i + 1] = transition[i].nextnum((double)forecastsemstud[i], true,true);
                i++;
            }
            //int examforecastsem = 0;
            forecastexam = examforecast(examforecastsem, getstud(examforecastsem), examtransition);
        }

        public programbatchclass(double?[] actualstud, int prog, string bstart) //real data
        {
            maxid++;
            this.id = maxid;
            this.progid = prog;
            this.actualbatch = true;
            this.batchstart = util.semester4to2(bstart);
            actualsemstud = actualstud;
            forecastsemstud[0] = null;
        }

        public programbatchclass(double?[] actualstud, int prog, string bstart, int exam, double?[] appl, int res) //real data
        {
            maxid++;
            this.id = maxid;
            this.progid = prog;
            this.actualbatch = true;
            this.batchstart = util.semester4to2(bstart);
            actualsemstud = actualstud;
            forecastsemstud[0] = null;
            actualexam = exam;
            Array.Copy(appl,this.applicants,4);
            reserves = res;
        }

        public void extrapolate(transitionclass[] transition, transitionclass[] examtransition)
        {
            int k = lastrealsemester();
            while (k >= 0 && transition[k] != null)
            {
                forecastsemstud[k + 1] = transition[k].nextnum(getstud(k), true, true);
                k++;
            }
            forecastexam = examforecast(examforecastsem, getstud(examforecastsem), examtransition);
        }

        public int lastrealsemester() //extend real data with forecast
        {
            int k = maxsem;
            do
                k--;
            while (k > 0 && actualsemstud[k] == null);
            return k;
        }

        
        public double? getactualstud(int sem)
        {
            return actualsemstud[sem] + exchangestud;
        }

        public double getstud(int sem) //termin i programmet; sem=1 => T1 etc.
        {
            if (sem < 0)
                return 0;
            if (actualsemstud.Length < sem+1)
                return 0;
            else if (actualsemstud[sem] != null)
                return (double)actualsemstud[sem]+exchangestud;
            else if (forecastsemstud[sem] != null)
                return (double)forecastsemstud[sem];
            else
                return 0;
        }

        public void setstud(double? stud,int sem)
        {
            if (sem < actualsemstud.Length)
                actualsemstud[sem] = stud;
        }

        public double getstud(string sem) //kalendertermin; sem = "VT21" etc.
        {
            string ss = batchstart;
            int isem = 1;
            while (isem < maxsem)
            {
                if (ss == sem)
                    return getstud(isem);
                isem++;
                ss = util.incrementsemester(ss);
            }
            return 0;
        }

        public double getmeanstud(int fromsem,int tosem)
        {
            double sum = 0;
            int nsem = 0;
            for (int i = fromsem; i <= tosem; i++)
            {
                sum += getstud(i);
            }
            return sum / (tosem - fromsem + 1);
        }

        public double getyearstud(int year)
        {
            int yr = year;
            if (yr > 2000)
                yr -= 2000;
            int startyear = util.semtoint(batchstart);
            if (yr < startyear)
                return 0;

            double stud = 0;
            int isem = 1;
            string sem = batchstart;
            while (isem < maxsem)
            {
                if (util.semtoint(sem) == yr)
                    stud += getstud(isem);
                sem = util.incrementsemester(sem);
                isem++;
            }
            return stud;
        }

        public double examforecast(int sem, double stud, transitionclass[] examtransition)
        {
            if (examtransition[sem] == null)
            {
                return 0;
            }
            else
                return examtransition[sem].nextnum(stud, true, true);
        }

    }
}
