﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace ProgramPrognos
{
    public class programclass
    {
        public static int maxid;
        public static int prodyear = 2020;
        public static int lastbatchyear = 2021;
        public static double defaultretention = 0.85; //per semester

        public string name = "";
        public List<string> coursecodelist = new List<string>();
        public List<string> applcodelist = new List<string>();
        public string subjectcode = "";
        public int id = -1;
        public int semesters = -1;
        public double hp = -1;
        public string area = ""; // L = lärarutb, V = vårdutb, T = Teknik500, blank = övrigt
        public string subject = "x";
        public string sector = "x";
        public string utype = "Program";
        public List<string> partofpackage = new List<string>();
        public static Dictionary<string, string> areanamedict = new Dictionary<string, string>()
        {
            {"L","Lärarutbildning" },
            {"V", "Vårdutbildning" },
            {"T","Teknik500" },
            {"","Övrigt" }
        };
        public string homeinst = Form1.utaninst;

        //Kullar och övergångar
        public List<programbatchclass> batchlist = new List<programbatchclass>(); //alla kullar, inklusive extrapolerade, både betalande och anslagsfinansierade
        public List<programbatchclass> payingbatchlist = new List<programbatchclass>(); //alla kullar, inklusive extrapolerade, enbart betalande
        public transitionclass[] transition = new transitionclass[programbatchclass.maxsem];   //transition from Tn to Tn+1
        public transitionclass[] examtransition = new transitionclass[programbatchclass.maxsem]; //transition from Tn to exam
        public Dictionary<int, transitionclass[]> appltransition = new Dictionary<int, transitionclass[]>(); //transition from applicants to Tn

        //Referensvärden:
        public fracprodclass totalprod = new fracprodclass(); //totalt referensåret
        public fracprodclass prod_per_student = new fracprodclass(); //per student referensåret
        public Dictionary<string, fracprodclass> fracproddict = new Dictionary<string, fracprodclass>(); //per institution per student

        //En simuleringsomgång:
        public Dictionary<int, fracprodclass> yearproddict = new Dictionary<int, fracprodclass>(); //per år total prod
        public Dictionary<int, Dictionary<string, fracprodclass>> yearinstproddict = new Dictionary<int, Dictionary<string, fracprodclass>>(); //per år per inst prod
        public Dictionary<int, double> examforecastdict = new Dictionary<int, double>(); //examina ¨per år

        //Summerat över simuleringsomgångar:
        public Dictionary<int, forecastrangeclass> yearprodrangedict = new Dictionary<int, forecastrangeclass>(); //per år total prod
        public Dictionary<int, Dictionary<string, forecastrangeclass>> yearinstprodrangedict = new Dictionary<int, Dictionary<string, forecastrangeclass>>(); //per år per inst prod
        public Dictionary<int, forecastrangeclass> examforecastrangedict = new Dictionary<int, forecastrangeclass>();

        //Kurser i programmet
           // från program_programkurser_intag:
        public Dictionary<int, Dictionary<string, double>> coursedict = new Dictionary<int, Dictionary<string, double>>(); // [termin][kursnamn][andel av stud]; kursnamn är index till fkdict
           // från Lokal ersättning programkurser:
        public Dictionary<string, double> coursehstdict = new Dictionary<string, double>();
        public Dictionary<string, double> courseincomedict = new Dictionary<string, double>();
        public Dictionary<string, double> coursecostdict = new Dictionary<string, double>();

        //diverse
        public double averageaccepted = 0;
        public bool fk = false; //fristående kurser som pseudo-program
        public double fee = 0; //payment for whole program, for paying students
        public Dictionary<string, double> studentpengarea = new Dictionary<string, double>();
        public bool activecourse = false; //either fed from program or with recent HST

        public programclass clone(bool keepstudents)
        {
            programclass pc = new programclass(this.name, this.semesters, this.area);
            pc.transition = this.transition;
            pc.examtransition = this.examtransition;
            pc.appltransition = transitionclass.clone(this.appltransition);
            pc.totalprod = this.totalprod.clone();
            pc.fracproddict.Add(Form1.hda, pc.totalprod.clone());
            pc.prod_per_student = this.prod_per_student;
            //pc.fracproddict = this.fracproddict;
            pc.fracproddict = new Dictionary<string, fracprodclass>();
            foreach (string inst in this.fracproddict.Keys)
                pc.fracproddict.Add(inst, this.fracproddict[inst].clone());
            //pc.yearproddict = this.yearproddict;
            pc.averageaccepted = keepstudents? this.averageaccepted : 0;
            pc.fk = this.fk;
            pc.homeinst = this.homeinst;
            pc.semesters = this.semesters;
            pc.hp = this.hp;
            pc.utype = this.utype;
            pc.fee = this.fee;
            pc.partofpackage = this.partofpackage;

            foreach (string cl in this.coursecodelist)
                pc.coursecodelist.Add(cl);
            foreach (string cl in this.applcodelist)
                pc.applcodelist.Add(cl);

            if (keepstudents)
            {
                foreach (programbatchclass bc in this.batchlist)
                {
                    if (bc.actualbatch)
                        pc.batchlist.Add(bc.cloneactual());
                }
            }

            return pc;

        }
        public programclass clone()
        {
            return clone(true);
        }

        public static programclass clone(List<programclass> qprog) //returns clone that is average of programs in list. No students.
        {
            programclass pc = qprog.First().clone(false); //basic data frpm first in list
            pc.name = "Nytt program";
            pc.coursecodelist.Clear();
            pc.applcodelist.Clear();
            pc.batchlist.Clear();
            pc.transition = transitionclass.average((from c in qprog select c.transition).ToList());
            pc.examtransition = transitionclass.average((from c in qprog select c.examtransition).ToList());
            for (int j = 0; j < 4; j++)
                pc.appltransition[j] = transitionclass.average((from c in qprog select c.appltransition[j]).ToList());
            pc.totalprod = fracprodclass.average((from c in qprog select c.totalprod).ToList());
            pc.prod_per_student = fracprodclass.average((from c in qprog select c.prod_per_student).ToList());
            pc.fracproddict.Clear();
            foreach (string inst in Form1.institutiondict.Keys)
            {
                var qi = from c in qprog where c.fracproddict.ContainsKey(inst) select c.fracproddict[inst];
                if (qi.Count() > 0)
                    pc.fracproddict.Add(inst, fracprodclass.average(qi.ToList()));
            }

            pc.fracproddict.Add(Form1.hda, pc.totalprod.clone());

            return pc;
        }

        public void datafromtemplate(programclass template)
        {
            this.transition = template.transition;
            this.examtransition = template.examtransition;
            this.appltransition = transitionclass.clone(template.appltransition);
            this.totalprod = template.totalprod.clone();
            this.fracproddict.Add(Form1.hda, this.totalprod.clone());
            this.prod_per_student = template.prod_per_student;
            this.fracproddict = new Dictionary<string, fracprodclass>();
            foreach (string inst in template.fracproddict.Keys)
                this.fracproddict.Add(inst, template.fracproddict[inst].clone());
            //this.yearproddict = template.yearproddict;
            this.averageaccepted = template.averageaccepted;
            this.fk = template.fk;
            this.homeinst = template.homeinst;
            if (this.semesters <= 0)
                this.semesters = template.semesters;
            if (this.hp <= 0)
                this.hp = template.hp;
            this.utype = template.utype;
            this.fee = template.fee;
            this.partofpackage = template.partofpackage;

        }

        public programclass(string namepar)
        {
            this.name = namepar;
            maxid++;
            this.id = maxid;
            if (this.name.StartsWith("FK "))
                this.fk = true;
            for (int i = 0; i < 4; i++)
                this.appltransition.Add(i, new transitionclass[programbatchclass.maxsem]);
        }

        public programclass(string namepar, int sem, string progarea)
        {
            this.name = namepar;
            maxid++;
            this.id = maxid;
            if (this.name.StartsWith("FK "))
                this.fk = true;
            this.semesters = sem;
            this.hp = semesters * 30;
            this.area = progarea;
            for (int i = 0; i < 4; i++)
                this.appltransition.Add(i, new transitionclass[programbatchclass.maxsem]);
        }

        public programclass(string namepar,int sem, double ret, string progarea)
        {
            this.name = namepar;
            maxid++;
            this.id = maxid;
            if (this.name.StartsWith("FK "))
                this.fk = true;
            this.semesters = sem;
            this.hp = semesters * 30;
            this.area = progarea;
            this.extendretention(ret);
            this.prod_per_student = generate_prod(30000, 20000, 0.8);
            this.fracproddict.Add(Form1.utaninst, prod_per_student.clone());
            this.fracproddict.Add(Form1.hda, this.totalprod.clone());
            for (int i = 0; i < 4; i++)
                this.appltransition.Add(i, new transitionclass[programbatchclass.maxsem]);
        }

        public fracprodclass generate_prod(double hstpeng,double hprpeng,double prestation)
        {
            fracprodclass fp = new fracprodclass();
            fp.frachst = 0.5;
            fp.frachpr = fp.frachst * prestation;
            fp.frachstmoney = fp.frachst * hstpeng;
            fp.frachprmoney = fp.frachpr * hprpeng;
            fp.fracmoney = fp.frachstmoney + fp.frachprmoney;
            return fp;
        }

        public double getstudents(int year)
        {
            double stud = 0;
            foreach (programbatchclass b in batchlist)
                stud += b.getyearstud(year);
            return stud;
        }

        public programbatchclass getbatch(string sem)
        {
            var q = from c in batchlist where c.batchstart == sem select c;
            if (q.Count() == 1)
                return q.First();
            else
                return null;
        }

        public programbatchclass getnextbatch(string sem)
        {
            string nextsem = sem;
            do
            {
                nextsem = util.incrementsemester(nextsem);
            }
            while ((getbatch(nextsem) == null) && (util.semtoint(nextsem) < 25));

            return getbatch(nextsem);
        }

        public programbatchclass getfirstbatch()
        {
            string minsem = "HT99";
            programbatchclass pbmin = null;
            foreach (programbatchclass pb in batchlist)
            {
                if (util.comparesemesters(pb.batchstart,minsem))
                {
                    minsem = pb.batchstart;
                    pbmin = pb;
                }
            }
            return pbmin;
        }

        public void add_production(string inst, double hst, double hpr, double hstkr, double hprkr, double kr)
        {
            Console.WriteLine(name + ": " + hst);
            if (!fracproddict.ContainsKey(inst))
                fracproddict.Add(inst, new fracprodclass());
            if (!fracproddict.ContainsKey(Form1.hda))
                fracproddict.Add(Form1.hda, new fracprodclass());
            totalprod.frachst += hst;
            totalprod.frachpr += hpr;
            totalprod.frachstmoney += hstkr;
            totalprod.frachprmoney += hprkr;
            totalprod.fracmoney += kr;
            totalprod.updatepeng();
            fracproddict[inst].frachst += hst;
            fracproddict[inst].frachpr += hpr;
            fracproddict[inst].frachstmoney += hstkr;
            fracproddict[inst].frachprmoney += hprkr;
            fracproddict[inst].fracmoney += kr;
            fracproddict[inst].updatepeng();
            fracproddict[Form1.hda].frachst += hst;
            fracproddict[Form1.hda].frachpr += hpr;
            fracproddict[Form1.hda].frachstmoney += hstkr;
            fracproddict[Form1.hda].frachprmoney += hprkr;
            fracproddict[Form1.hda].fracmoney += kr;
            fracproddict[Form1.hda].updatepeng();
        }

        public void fill_transition_gaps()
        {
            if (homeinst == Form1.utaninst)
                return;

            for (int i=0;i<=semesters;i++)
            {
                if (transition[i] == null)
                    transition[i] = Form1.institutiondict[homeinst].meantransition[i];
                if (examtransition[i] == null)
                    examtransition[i] = Form1.institutiondict[homeinst].meanexamtransition[i];
            }
        }

        public void calculate_transitions()
        {
            //double defaultprob = 0.9;
            //double defaultexamprob = 0.8;
            Dictionary<int, double[]> tdict = new Dictionary<int, double[]>();
            Dictionary<int, double[]> tdictexam = new Dictionary<int, double[]>();
            Dictionary<int, Dictionary<int, double[]>> tdictappl = new Dictionary<int, Dictionary<int, double[]>>();
            int nacc = 0;
            double sumacc = 0;
            foreach (programbatchclass bc in batchlist)
            {
                for (int i=0;i<programbatchclass.maxsem-1;i++)
                {
                    //terminsövergångar:
                    if ((bc.getstud(i) > 0) && (bc.getstud(i+1) > 0))
                    {
                        if (tdict.ContainsKey(i))
                        {
                            tdict[i][0] += bc.getstud(i);
                            tdict[i][1] += bc.getstud(i+1);
                        }
                        else
                        {
                            tdict.Add(i, new double[] { bc.getstud(i), bc.getstud(i + 1) });
                        }
                    }
                    //från termin till examen:
                    if ( bc.actualexam != null && bc.getstud(this.semesters) > 0) //har kullen hunnit till sista terminen?
                    {
                        if (tdictexam.ContainsKey(i))
                        {
                            tdictexam[i][0] += bc.getstud(i);
                            tdictexam[i][1] += (double)bc.actualexam;
                        }
                        else
                        {
                            tdictexam.Add(i, new double[] { bc.getstud(i), (double)bc.actualexam });
                        }
                    }
                    //från sökande till termin:
                    for (int j = 0; j < 4; j++)
                    {
                        if (!tdictappl.ContainsKey(j))
                            tdictappl.Add(j, new Dictionary<int, double[]>());
                        if (bc.applicants[j] > 0)
                        {
                            if (tdictappl[j].ContainsKey(i))
                            {
                                tdictappl[j][i][0] += bc.getstud(i);
                                tdictappl[j][i][1] += (double)bc.applicants[j];
                            }
                            else
                            {
                                tdictappl[j].Add(i, new double[] { bc.getstud(i), (double)bc.applicants[j] });
                            }
                        }
                    }
                }
                if (bc.getstud(0) > 0 && (lastbatchyear%2000-util.semtoint(bc.batchstart) < 4))
                {
                    nacc++;
                    sumacc += bc.getstud(0);
                }
            }
            int imax = -1;
            for (int i = 0; i < programbatchclass.maxsem - 1; i++)
            {
                if (tdict.ContainsKey(i))
                {
                    double tprob = tdict[i][1] / tdict[i][0];
                    transition[i] = new transitionclass(tprob, Math.Sqrt(tprob));
                    imax = i;
                }
                //else if (i < this.semesters)
                //{
                //    transition[i] = new transitionclass(defaultprob,Math.Sqrt(defaultprob));
                //}
                else
                {
                    transition[i] = null;
                    //transition[i] = new transitionclass(defaultprob,Math.Sqrt(defaultprob));
                }
                if (i <= this.semesters)
                {
                    if (tdictexam.ContainsKey(i))
                    {
                        double tprob = tdictexam[i][1] / tdictexam[i][0];
                        examtransition[i] = new transitionclass(tprob, Math.Sqrt(tprob));
                        imax = i;
                    }
                    else
                        examtransition[i] = null;// new transitionclass(defaultexamprob, Math.Sqrt(defaultexamprob));
                }
                if (i <= this.semesters)
                {
                    for (int j = 0; j < 4; j++)
                    {
                        if (!appltransition.ContainsKey(j))
                            appltransition.Add(j,new transitionclass[programbatchclass.maxsem]);
                        if (tdictappl.ContainsKey(j) && tdictappl[j].ContainsKey(i))
                        {
                            double tprob = tdictappl[j][i][0] / tdictappl[j][i][1];
                            appltransition[j][i] = new transitionclass(tprob, Math.Sqrt(tprob));
                            imax = i;
                        }
                        else
                            appltransition[j][i] = null;// new transitionclass(defaultapplprob, Math.Sqrt(defaultapplprob));
                    }
                }
            }
            //this.semesters = imax+1;

            if (nacc > 0)
                averageaccepted = sumacc / nacc;

        }

        public double averageretention()
        {
            int nt = 0;
            double tsum = 0;
            for (int i = 0; i < programbatchclass.maxsem - 1; i++)
            {
                if (transition[i] != null)
                {
                    nt++;
                    tsum += transition[i].transitionprob;
                }
            }
            if (nt > 0)
                return tsum / nt;
            else
                return -1;

        }

        public void replaceretention(double ret)
        {
            for (int i = 0; i < programbatchclass.maxsem - 1; i++)
            {
                if (transition[i] != null)
                {
                    transition[i].transitionprob = ret;
                }
            }
        }

        public void extendretention(double ret)
        {
            for (int i = 0; i < semesters; i++)
            {
                if (transition[i] == null)
                {
                    transition[i] = new transitionclass(ret, Math.Sqrt(ret));
                }
            }
        }

        public void normalize_per_student()
        {
            double nstud = getstudents(prodyear);
            if (nstud == 0)
                nstud = totalprod.frachst;
            if (nstud == 0)
                nstud = totalprod.frachpr;
            Console.WriteLine(name+" nstud = " + nstud);
            this.prod_per_student = this.totalprod.clone();
            this.prod_per_student.normalize(nstud);
            foreach (string inst in fracproddict.Keys)
                fracproddict[inst].normalize(nstud);
        }

        public void set_homeinst()
        {
            if (homeinst != Form1.utaninst)
                return;

            double pmax = 0;
            foreach (string inst in fracproddict.Keys)
            {
                if (inst == Form1.hda)
                    continue;
                if (fracproddict[inst].fracmoney > pmax)
                {
                    homeinst = inst;
                    pmax = fracproddict[inst].fracmoney;
                }
            }

            if (homeinst == Form1.utaninst && this.coursedict.Count > 0)
            {
                Dictionary<string, double> studdict = new Dictionary<string, double>();
                Dictionary<string, double> ccdict = new Dictionary<string, double>();
                double totstud = 0;
                studdict.Add(this.homeinst, 0);
                foreach (int sem in this.coursedict.Keys)
                {
                    foreach (string code in this.coursedict[sem].Keys)
                    {
                        totstud += this.coursedict[sem][code] * Form1.fkcodedict[code].hp;
                        string cc = Form1.fkcodedict[code].homeinst;
                        if (studdict.ContainsKey(cc))
                            studdict[cc] += this.coursedict[sem][code] * Form1.fkcodedict[code].hp;
                        else
                            studdict.Add(cc, this.coursedict[sem][code] * Form1.fkcodedict[code].hp);
                        if (ccdict.ContainsKey(cc))
                            ccdict[cc]++;
                        else
                            ccdict.Add(cc, 1);
                    }
                }
                if (totstud > 0)
                {
                    double maxstud = 0;
                    foreach (string inst in Form1.instshortdict.Keys)
                    {
                        if (studdict.ContainsKey(inst) && studdict[inst] > maxstud)
                        {
                            maxstud = studdict[inst];
                            this.homeinst = inst;
                        }
                    }
                }
                else
                {
                    double ccmax = 0;
                    foreach (string inst in Form1.instshortdict.Keys)
                    {
                        if (ccdict.ContainsKey(inst) && ccdict[inst] > ccmax)
                        {
                            ccmax = ccdict[inst];
                            this.homeinst = inst;
                        }
                    }
                }

            }

            if (homeinst == Form1.utaninst && name == "Produktionstekniker 120 hp")
                homeinst = "Institutionen för information och teknik";
        }

        public Dictionary<string, double> extrapolate(bool futureadm, int endyear)
        {
            return extrapolate(lastbatchyear, endyear,futureadm);
        }

        public Dictionary<string, double> extrapolate(Dictionary<string,double> plstuddict, int endyear, bool futureadm)
        {
            int baseyear = lastbatchyear;
            //int endyear = lastbatchyear + 5;
            return extrapolate(baseyear, endyear, plstuddict,futureadm);
        }

        public Dictionary<string, double> extrapolate(int baseyear, int endyear, bool futureadm)
        {
            return extrapolate(baseyear, endyear, new Dictionary<string, double>(), futureadm);
            //if (fk)
            //    return extrapolate_fk(baseyear, endyear, new Dictionary<string,double>());
            //else
            //    return extrapolate_program(baseyear, endyear, new Dictionary<string, double>());
        }

        public Dictionary<string, double> extrapolate(int baseyear, int endyear, Dictionary<string, double> plstuddict, bool futureadm)
        {
            Dictionary<string, double> xdict;
            if (fk)
                xdict = extrapolate_fk(baseyear, endyear, plstuddict,futureadm);
            else
                xdict = extrapolate_program(baseyear, endyear, plstuddict,futureadm);

            addtorange();

            return xdict;
        }

        private Dictionary<string, double> extrapolate_program(int baseyear, int endyear, Dictionary<string, double> plstuddict, bool futureadm)
        {
            List<string> oldbatches = new List<string>();
            bool htstart = false;
            bool vtstart = false;
            int lastbatch = -1;
            foreach (programbatchclass bc in batchlist)
            {
                oldbatches.Add(bc.batchstart);
                if (util.semtoint(bc.batchstart) == lastbatchyear % 2000)
                {
                    if (bc.batchstart.StartsWith("HT"))
                        htstart = true;
                    else if (bc.batchstart.StartsWith("VT"))
                        vtstart = true;
                    if (util.semtoint(bc.batchstart) > lastbatch)
                        lastbatch = util.semtoint(bc.batchstart);
                }
                bc.extrapolate(transition,examtransition);
            }

            if (futureadm && (lastbatch >= lastbatchyear % 2000 || Form1.scenarioloaded)) //did we actually accept students in 2021 or later? Otherwise skip new recruitment.
            {
                bool frompsd = plstuddict.Count > 0;
                string newbatch = "VT" + baseyear%2000;
                while (util.semtoint(newbatch) <= endyear % 2000)
                {
                    if (!oldbatches.Contains(newbatch))
                    {
                        if (frompsd || Form1.scenarioloaded)
                        {
                            if (plstuddict.ContainsKey(newbatch))
                            {
                                programbatchclass bc = new programbatchclass(plstuddict[newbatch], this.id, newbatch, transition,examtransition);
                                batchlist.Add(bc);
                            }
                        }
                        else //generate new batches from historical accept-numbers
                        {
                            if ((htstart && newbatch.StartsWith("HT")) || (vtstart && newbatch.StartsWith("VT")))
                            {
                                double accepted = averageaccepted;
                                plstuddict.Add(newbatch, accepted);
                                programbatchclass bc = new programbatchclass(accepted, this.id, newbatch, transition,examtransition);
                                batchlist.Add(bc);
                            }
                        }
                    }
                    newbatch = util.incrementsemester(newbatch);
                }
            }

            for (int year = baseyear; year <= endyear; year++)
                sum_production(year);

            return plstuddict;
        }


        private Dictionary<string, double> extrapolate_fk(int baseyear, int endyear, Dictionary<string, double> plstuddict, bool futureadm)
        {
            if (!futureadm)
                return plstuddict;

            if (plstuddict.Count == 0)
            {
                for (int year = baseyear; year <= endyear; year++)
                {
                    yearproddict.Add(year, totalprod.clone());
                    yearinstproddict.Add(year, new Dictionary<string, fracprodclass>());
                    foreach (string inst in fracproddict.Keys)
                    {
                        yearinstproddict[year].Add(inst, fracproddict[inst].clone());
                        yearinstproddict[year][inst].normalize(1 / totalprod.frachst);
                    }
                    plstuddict.Add("VT" + year%2000, 0.5*yearproddict[year].frachst);
                    plstuddict.Add("HT" + year%2000, 0.5*yearproddict[year].frachst);
                }
            }
            else
            {
                Dictionary<int, double> yearhst = new Dictionary<int, double>();
                foreach (string sem in plstuddict.Keys)
                {
                    int yy = util.semtoint(sem);
                    if (!yearhst.ContainsKey(yy))
                        yearhst.Add(yy, plstuddict[sem]);
                    else
                        yearhst[yy] += plstuddict[sem];
                }

                for (int year = baseyear; year <= endyear; year++)
                {
                    yearproddict.Add(year, prod_per_student.clone());
                    yearproddict[year].normalize(1 / yearhst[year % 2000]);
                    yearinstproddict.Add(year, new Dictionary<string, fracprodclass>());
                    foreach (string inst in fracproddict.Keys)
                    {
                        yearinstproddict[year].Add(inst, fracproddict[inst].clone());
                        yearinstproddict[year][inst].normalize(1 / yearproddict[year].frachst);
                    }
                }
            }

            return plstuddict;
        }

        public void sum_production(int year)
        {
            fracprodclass yt = new fracprodclass();
            Dictionary<string, fracprodclass> yidict = new Dictionary<string, fracprodclass>();
            foreach (string inst in fracproddict.Keys)
                yidict.Add(inst, new fracprodclass());

            foreach (programbatchclass bc in batchlist)
            {
                double nstud = bc.getyearstud(year);
                yt.sumstudents(nstud, prod_per_student);
                foreach (string inst in fracproddict.Keys)
                    yidict[inst].sumstudents(nstud, fracproddict[inst]);
                if (bc.forecastexam != null)
                {
                    int finalyear = util.year2to4(util.semtoint(util.shiftsemester(bc.batchstart, semesters)));
                    if (finalyear == year)
                    {
                        if (!examforecastdict.ContainsKey(finalyear))
                            examforecastdict.Add(finalyear, (double)bc.forecastexam);
                        else
                            examforecastdict[finalyear] += (double)bc.forecastexam;
                    }
                }

            }

            yearproddict.Add(year, yt);
            yearinstproddict.Add(year, yidict);

        }

        public void addtorange() //adds one simulation round to range dicts
        {
            Console.WriteLine("addtorange " + this.name);
            foreach (int year in yearproddict.Keys)
            {
                if (!yearprodrangedict.ContainsKey(year))
                    yearprodrangedict.Add(year, new forecastrangeclass());
                yearprodrangedict[year].Add(yearproddict[year].fracmoney);
                yearprodrangedict[year].Add(yearproddict[year]);
            }

            foreach (int year in yearinstproddict.Keys)
            {
                if (!yearinstprodrangedict.ContainsKey(year))
                    yearinstprodrangedict.Add(year, new Dictionary<string, forecastrangeclass>());
                foreach (string inst in yearinstproddict[year].Keys)
                {
                    if (!yearinstprodrangedict[year].ContainsKey(inst))
                    {
                        yearinstprodrangedict[year].Add(inst, new forecastrangeclass());
                    }
                    yearinstprodrangedict[year][inst].Add(yearinstproddict[year][inst].fracmoney);
                    yearinstprodrangedict[year][inst].Add(yearinstproddict[year][inst]);
                }
            }

            foreach (int year in examforecastdict.Keys)
            {
                if (!examforecastrangedict.ContainsKey(year))
                    examforecastrangedict.Add(year, new forecastrangeclass());
                examforecastrangedict[year].Add(examforecastdict[year]);
            }
        }

        public double[] batchsemsum(string beforebatch)
        {
            double[] tstud = new double[programbatchclass.maxsem];
            for (int i = 0; i < programbatchclass.maxsem; i++)
                tstud[i] = 0;
            programbatchclass pb = null;
            string bstart = beforebatch;
            do
            {
                pb = this.getnextbatch(bstart);
                if (pb != null)
                {
                    for (int i = 1; i <= this.semesters; i++)
                    {
                        double? xstud = pb.getactualstud(i);
                        if (xstud > 0)
                            tstud[i] += (double)xstud;
                    }
                    bstart = pb.batchstart;
                }
            }
            while (pb != null);
            return tstud;
        }


        public int examsum(int startyear,int endyear)
        {
            double sum = 0;
            for (int year=startyear;year <= endyear;year++)
            {
                if (examforecastrangedict.ContainsKey(year))
                    sum += examforecastrangedict[year].Average();
            }
            return (int)Math.Round(sum);
        }

        public bool is_advanced()
        {
            if (semesters > 4)
                return false;
            else if (name.ToLower().Contains("master"))
                return true;
            else if (name.ToLower().Contains("magister"))
                return true;
            return false;
        }

        string rex1 = @"[GA]([\p{Lu}][\p{Lu}])\w+";
        string rex2 = @"([\p{Lu}][\p{Lu}])\w+";
        public string bestcode()
        {
            foreach (string code in coursecodelist)
            {
                foreach (Match m in Regex.Matches(code, rex1))
                {
                    return code;
                }
            }
            if (coursecodelist.Count > 0)
                return coursecodelist.First();
            return "";
        }

        public int getallT1(string startsem,string endsem)
        {
            double n = 0;
            string sem = startsem;
            while (sem != endsem)
            {
                var bc = getbatch(sem);
                if (bc != null)
                    n += bc.getstud(1);
                sem = util.incrementsemester(sem);
            }
            return (int)n;
        }

        public bool hasbeginners(string startsem,string endsem)
        {
            return getallT1(startsem, endsem) > 0;
        }
    }

}
