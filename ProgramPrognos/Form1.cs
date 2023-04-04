using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;

namespace ProgramPrognos
{
    public partial class Form1 : Form
    {
        //public static string homefolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        public static string homefolder = Environment.GetEnvironmentVariable("onedrive")+@"\Dokument";
        public static string folder = homefolder + @"\Invärld\Långtidsbudget";
        public static string docfolder = homefolder;//@"\\dustaff\home\"+Environment.UserName+@"\Documents";
        public static string utaninst = "Utan institution";

        public static List<string> sectorlist = new List<string>();


        public static Dictionary<string, string> programkoppling = new Dictionary<string, string>();
        public static Dictionary<string, programclass> programdict = new Dictionary<string, programclass>();
        public static Dictionary<string, programclass> origprogramdict = new Dictionary<string, programclass>();
        public static Dictionary<string, institutionclass> institutiondict = new Dictionary<string, institutionclass>();
        public static Dictionary<string, string> instshortdict = new Dictionary<string, string>();
        public static Dictionary<string, string> shortinstdict = new Dictionary<string, string>();
        public static Dictionary<string, string> subjinstdict = new Dictionary<string, string>(); // from coursecode subjects to institutions
        public static bool scenarioloaded = false;

        public static Dictionary<string,programclass> fkdict = new Dictionary<string,programclass>();
        public static Dictionary<string, programclass> fkcodedict = new Dictionary<string, programclass>();
        public static Dictionary<string, programclass> paketdict = new Dictionary<string, programclass>();

        public static Dictionary<string, double> lokal_ers_hst = new Dictionary<string, double>();
        public static Dictionary<string, double> lokal_ers_hpr = new Dictionary<string, double>();

        public int endyear = -1;

        public Dictionary<string, Dictionary<string, double>> plannedstudents = new Dictionary<string, Dictionary<string, double>>();
        public Form1()
        {
            InitializeComponent();
            string machine = Environment.MachineName;
            var df = homefolder;//Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            memo("Machine name = " + Environment.MachineName);

            datafolderlabel.Text = "Data folder: " + folder;
            docfolderlabel.Text = "Document folder " + docfolder;

            string fn = folder + @"\programkopplingar.txt";
            string[] ss = Directory.GetDirectories(homefolder);
            Directory.GetFiles(folder);

            using (StreamReader sr = new StreamReader(fn))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    string[] words = line.Split('\t');
                    if (words.Length < 2)
                        continue;
                    if (programkoppling.ContainsKey(words[0]))
                        continue;
                    if (words[1].StartsWith("#"))
                        words[1] = words[0];
                    programkoppling.Add(words[0], words[1]);
                    int sem = 6;
                    if (words.Length >= 3)
                        sem = util.tryconvert(words[2]);
                    string progarea = "";
                    if (words.Length >= 4)
                        progarea = words[3];
                    if (!origprogramdict.ContainsKey(words[1]))
                        origprogramdict.Add(words[1], new programclass(words[1], sem, progarea));
                    origprogramdict[words[1]].applcodelist.Add(words[0]);
                }
            }
            instshortdict.Add("Institutionen för hälsa och välfärd", "HV");
            instshortdict.Add("Institutionen för information och teknik", "IT");
            instshortdict.Add("Institutionen för kultur och samhälle", "KS");
            instshortdict.Add("Institutionen för språk, litteratur och lärande", "SLS");
            instshortdict.Add("Institutionen för lärarutbildning", "LU");
            instshortdict.Add(utaninst, "?");
            foreach (string inst in instshortdict.Keys)
                shortinstdict.Add(instshortdict[inst], inst);

            fill_subjinstdict();

            read_studentpeng();
        }

        public void read_studentpeng()
        {
            int n = 0;
            string fn = folder + "\\ers_belopp_lokala 2021.txt";
            using (StreamReader sr = new StreamReader(fn))
            {
                sr.ReadLine();
                sr.ReadLine();
                sr.ReadLine();
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    string[] words = line.Split('\t');
                    string area = words[0].Substring(words[0].IndexOf('(') + 1, 2);
                    double hstpeng = util.tryconvertdouble(words[1].Replace(" ", ""));
                    double hprpeng = util.tryconvertdouble(words[2].Replace(" ", ""));
                    memo(area + "\t" + hstpeng + "\t" + hprpeng);
                    lokal_ers_hst.Add(area, hstpeng);
                    lokal_ers_hpr.Add(area, hprpeng);
                    n++;
                }
            }

            memo(n + " utbildningsområden i readstudentpeng");

        }

        public void fill_subjinstdict()
        {
            subjinstdict.Add("AB", "KS");
            subjinstdict.Add("AR", "SLS");
            subjinstdict.Add("AS", "KS");
            subjinstdict.Add("AU", "KS");
            subjinstdict.Add("BE", "?");
            subjinstdict.Add("BI", "LU");
            subjinstdict.Add("BP", "LU");
            subjinstdict.Add("BQ", "KS");
            subjinstdict.Add("BY", "IT");
            subjinstdict.Add("DT", "IT");
            subjinstdict.Add("EG", "IT");
            subjinstdict.Add("EN", "SLS");
            subjinstdict.Add("ET", "IT");
            subjinstdict.Add("EU", "KS");
            subjinstdict.Add("FI", "KS");
            subjinstdict.Add("FÖ", "KS");
            subjinstdict.Add("FR", "SLS");
            subjinstdict.Add("FY", "IT");
            subjinstdict.Add("GG", "LU");
            subjinstdict.Add("GT", "IT");
            subjinstdict.Add("HI", "KS");
            subjinstdict.Add("IE", "IT");
            subjinstdict.Add("IH", "HV");
            subjinstdict.Add("IK", "IT");
            subjinstdict.Add("IT", "SLS");
            subjinstdict.Add("JP", "SLS");
            subjinstdict.Add("KE", "IT");
            subjinstdict.Add("KG", "KS");
            subjinstdict.Add("KI", "SLS");
            subjinstdict.Add("KT", "?");
            subjinstdict.Add("LI", "SLS");
            subjinstdict.Add("LP", "KS");
            subjinstdict.Add("MA", "IT");
            subjinstdict.Add("MC", "HV");
            subjinstdict.Add("MD", "LU");
            subjinstdict.Add("MI", "IT");
            subjinstdict.Add("MÖ", "IT");
            subjinstdict.Add("MP", "IT");
            subjinstdict.Add("MT", "IT");
            subjinstdict.Add("NA", "KS");
            subjinstdict.Add("NV", "LU");
            subjinstdict.Add("PA", "KS");
            subjinstdict.Add("PE", "LU");
            subjinstdict.Add("PG", "LU");
            subjinstdict.Add("PR", "SLS");
            subjinstdict.Add("PS", "HV");
            subjinstdict.Add("RK", "KS");
            subjinstdict.Add("RV", "KS");
            subjinstdict.Add("RY", "SLS");
            subjinstdict.Add("SA", "HV");
            subjinstdict.Add("SK", "KS");
            subjinstdict.Add("SO", "KS");
            subjinstdict.Add("SP", "SLS");
            subjinstdict.Add("SQ", "KS");
            subjinstdict.Add("SR", "HV");
            subjinstdict.Add("SS", "SLS");
            subjinstdict.Add("ST", "IT");
            subjinstdict.Add("SV", "SLS");
            subjinstdict.Add("SW", "IT");
            subjinstdict.Add("TR", "KS");
            subjinstdict.Add("TY", "SLS");
            subjinstdict.Add("VÅ", "HV");
            subjinstdict.Add("VV", "HV");


        }

        public void memo(string s)
        {
            richTextBox1.AppendText(s + "\n");
            richTextBox1.ScrollToCaret();
        }


        private void Quitbutton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void read_retention(string fn)
        {
            int lbyear = -1;
            int nline = 0;
            using (StreamReader sr = new StreamReader(fn))
            {
                sr.ReadLine();//throw away two header lines
                sr.ReadLine();
                int offset = 2;


                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    string[] words = line.Split('\t');
                    string prog = words[0];
                    if (!origprogramdict.ContainsKey(prog))
                    {
                        if (programkoppling.ContainsKey(prog))
                            prog = programkoppling[prog];
                    }
                    if (!origprogramdict.ContainsKey(prog))
                    {
                        memo("Missing program " + prog);
                        continue;
                    }
                    double?[] actualstud = new double?[programbatchclass.maxsem];
                    for (int i = 0; i < programbatchclass.maxsem; i++)
                    {
                        if (offset + i >= words.Length)
                            actualstud[i] = null;
                        else
                            actualstud[i] = util.tryconvertnull(words[offset + i]);
                    }
                    programbatchclass b = new programbatchclass(actualstud, origprogramdict[prog].id, words[1]);
                    if (words[1].StartsWith("HT"))
                    {
                        int year = util.semtoint(words[1]);
                        if (year > lbyear)
                            lbyear = year;
                    }
                    origprogramdict[prog].batchlist.Add(b);
                    nline++;
                }
            }
            programclass.lastbatchyear = util.year2to4(lbyear);
            memo("Lastbatchyear = " + programclass.lastbatchyear);
            memo("nline = " + nline);
        }

        private void read_retention_v2(string fn)
        {
            int nline = 0;
            using (StreamReader sr = new StreamReader(fn))
            {
                sr.ReadLine();//throw away two header lines
                sr.ReadLine();
                int offset = 4;

                int progcol = 1;
                int batchcol = 0;
                int examcol = 2;
                int applcol = 3;
                int subjcol = -1;
                int sectorcol = -1;
                int reservecol = -1;

                if (fn.Contains("_classified"))
                {
                    subjcol = 0;
                    sectorcol = 1;
                    progcol = 3;
                    batchcol = 2;
                    examcol = 4;
                    applcol = 5;
                    reservecol = 6;
                    offset = 7;
                }

                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    string[] words = line.Split('\t');
                    string prog = words[progcol];
                    if (!origprogramdict.ContainsKey(prog))
                    {
                        if (programkoppling.ContainsKey(prog))
                            prog = programkoppling[prog];
                    }
                    if (!origprogramdict.ContainsKey(prog))
                    {
                        memo("Missing program " + prog);
                        continue;
                    }
                    double?[] actualstud = new double?[programbatchclass.maxsem];
                    bool anynotnull = false;
                    for (int i = 0; i < programbatchclass.maxsem; i++)
                    {
                        if (offset + i >= words.Length)
                            actualstud[i] = null;
                        else
                        {
                            actualstud[i] = util.tryconvertnull(words[offset + i]);
                            if (actualstud[i] != null)
                                anynotnull = true;
                        }
                    }
                    if (!anynotnull)
                        continue;
                    int exam = util.tryconvert(words[examcol]);
                    if (exam < 0)
                        exam = 0;
                    int appl = util.tryconvert(words[applcol]);
                    if (appl < 0)
                        appl = 0;
                    int res = 0;
                    if (reservecol >= 0)
                    {
                        res = util.tryconvert(words[reservecol]);
                        if (res < 0)
                            res = 0;
                    }
                    if (subjcol >= 0)
                    {
                        origprogramdict[prog].subject = words[subjcol];
                        origprogramdict[prog].sector = words[sectorcol];
                        if (!sectorlist.Contains(words[sectorcol]))
                            sectorlist.Add(words[sectorcol]);
                    }

                    programbatchclass b = new programbatchclass(actualstud, origprogramdict[prog].id, words[batchcol], exam, appl, res);
                    origprogramdict[prog].batchlist.Add(b);
                    nline++;
                }
            }
            memo("nline = " + nline);
        }

        private void read_prod(string fn)
        {
            memo("Reading production from " + fn);

            int nline = 0;
            double krsum = 0;
            using (StreamReader sr = new StreamReader(fn))
            {
                sr.ReadLine();//throw away two header lines
                sr.ReadLine();
                //int offset = 2;

                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    string[] words = line.Split('\t');
                    if (words.Length < 7)
                        continue;
                    string inst = words[0];
                    if (!institutiondict.ContainsKey(inst))
                        institutiondict.Add(inst, new institutionclass(inst));
                    string prog = programkoppling[words[1]];

                    double hst = util.tryconvertdouble(words[2]);
                    double hpr = util.tryconvertdouble(words[3]);
                    double hstkr = util.tryconvertdouble(words[4]);
                    double hprkr = util.tryconvertdouble(words[5]);
                    double kr = util.tryconvertdouble(words[6]);
                    krsum += kr;
                    origprogramdict[prog].add_production(inst, hst, hpr, hstkr, hprkr, kr);
                    institutiondict[inst].addbaseyearproduction(hst, hpr, hstkr, hprkr, kr);

                    nline++;
                }
            }
            memo("nline = " + nline);
            memo("krsum = " + krsum);
        }

        private void readdatabutton_Click(object sender, EventArgs e)
        {
            //string fn1 = folder + @"\programretention 220302_v2.txt";
            //tring fn1 = folder + @"\programretention 221010_v2_classified.txt";
            string fn1 = folder + @"\programretention 230208.txt";
            //if (fn1.Contains("_v2"))
                read_retention_v2(fn1);
            //else
            //    read_retention(fn1);
            foreach (string p in origprogramdict.Keys)
                origprogramdict[p].calculate_transitions();

            string fn2 = folder + @"\prod per inst och prog_YYYY.txt";
            string fnyear = fn2;
            int year = DateTime.Now.Year + 1;
            do
            {
                year--;
                fnyear = fn2.Replace("YYYY", year.ToString());
            }
            while (!File.Exists(fnyear));

            programclass.prodyear = year;
            read_prod(fnyear);


            foreach (string p in origprogramdict.Keys)
            {
                origprogramdict[p].normalize_per_student();
                memo(p + "\t" + origprogramdict[p].prod_per_student.frachst.ToString("N1") + " hst per stud");
                origprogramdict[p].set_homeinst();
            }

            foreach (string inst in institutiondict.Keys)
            {
                institutiondict[inst].calculate_meantransition();
            }
            foreach (string p in origprogramdict.Keys)
                origprogramdict[p].fill_transition_gaps();

            businessbutton.Enabled = true;
        }

        List<string> make_scenario_semesters(int startyear, int endyear)
        {
            List<string> semlist = new List<string>();
            string sem = "VT" + ((startyear) % 2000);
            while (util.semtoint(sem) <= endyear % 2000)
            {
                semlist.Add(sem);
                sem = util.incrementsemester(sem);
            }
            return semlist;
        }

        private string semheader(List<string> semlist)
        {

            StringBuilder sbhead = new StringBuilder("\tTerminer\tRetention\tInstitution\tProgramtyp");
            foreach (string sem in semlist)
            {
                sbhead.Append("\t" + sem);
            }
            return sbhead.ToString();
        }

        private void reset_programdict()
        {
            //programdict.Clear();
            if (programdict.Count > 0)
            {
                foreach (KeyValuePair<string, programclass> pcpair in origprogramdict)
                {
                    programclass pc = programdict[pcpair.Key];
                    programdict[pcpair.Key] = pcpair.Value.clone();

                    programdict[pcpair.Key].examforecastrangedict = pc.examforecastrangedict;
                    programdict[pcpair.Key].yearprodrangedict = pc.yearprodrangedict;
                    programdict[pcpair.Key].yearinstprodrangedict = pc.yearinstprodrangedict;
                }
            }
            else
            {
                foreach (KeyValuePair<string, programclass> pcpair in origprogramdict)
                    programdict.Add(pcpair.Key, pcpair.Value.clone());
            }

        }

        private void extrapolation_round(int iround)
        {
            reset_programdict();

            List<string> semlist = make_scenario_semesters(programclass.lastbatchyear, endyear);
            if (iround == 0)
            {
                memo(semheader(semlist) + "\t" + semheader(semlist));
                memo("programdict.Count " + programdict.Count);
            }

            foreach (programclass pc in programdict.Values)
            {
                //memo(pc.name);
                if (plannedstudents.ContainsKey(pc.name))
                {
                    plannedstudents[pc.name] = pc.extrapolate(plannedstudents[pc.name], endyear, CBfutureadm.Checked);
                }
                else if (!scenarioloaded)
                {
                    plannedstudents.Add(pc.name, pc.extrapolate(CBfutureadm.Checked, endyear));
                }
                else
                {
                    plannedstudents.Add(pc.name, pc.extrapolate(new Dictionary<string, double>(), endyear, CBfutureadm.Checked)); //call with empty dummy planning
                }
                //Dictionary<string, double> plstuddict = plannedstudents[pc.name];

                StringBuilder sb = new StringBuilder(pc.name);
                foreach (string semsem in semlist)
                {
                    if (plannedstudents[pc.name].ContainsKey(semsem))
                        sb.Append("\t" + plannedstudents[pc.name][semsem].ToString("N0"));
                    else
                        sb.Append("\t");
                }
                if (!pc.fk)
                {
                    sb.Append("\tExamina:");
                    foreach (string semsem in semlist)
                    {
                        string startsem = util.shiftsemester(semsem, -pc.semesters);
                        programbatchclass bc = pc.getbatch(startsem);
                        if (bc != null && bc.forecastexam != null)
                            sb.Append("\t" + ((double)(bc.forecastexam)).ToString("N0"));
                        else
                            sb.Append("\t");
                    }
                }
                if (iround == 0)
                    memo(sb.ToString());
            }

        }

        private void businessbutton_Click(object sender, EventArgs e)
        {
            int nround = util.tryconvert(TBxrounds.Text);

            if (endyear < 0)
                endyear = util.tryconvert(TB_endyear.Text);

            for (int iround = 0; iround < nround; iround++)
            {
                memo("============== " + iround + " =====================");
                Console.WriteLine("============== " + iround + " =====================");
                extrapolation_round(iround);
            }

            memo("Done extrapolating");
            savescenariobutton.Enabled = true;
            proddisplaybutton.Enabled = true;
            businessbutton.Enabled = false;
            buttonExamforecast.Enabled = true;
        }

        private string formatmillions(double sum)
        {
            return (sum / 1e6).ToString("N1");
        }

        private void button1_Click(object sender, EventArgs e) //proddisplaybutton
        {
            int baseyear = programclass.lastbatchyear;
            if (endyear < 0)
                endyear = util.tryconvert(TB_endyear.Text);
            //int endyear = baseyear + 5;
            StringBuilder sbhead = new StringBuilder("");
            StringBuilder sblonghead = new StringBuilder("");
            StringBuilder sbbase = new StringBuilder("Base year");
            double basesum = 0;
            foreach (string inst in institutiondict.Keys)
            {
                institutiondict[inst].clearproduction(baseyear, endyear);
                sbhead.Append("\t" + instshortdict[inst]);
                sblonghead.Append("\t\t\t" + instshortdict[inst] + "\t");
                sbbase.Append("\t" + formatmillions(institutiondict[inst].baseyearprod.fracmoney));
                basesum += institutiondict[inst].baseyearprod.fracmoney;
            }
            sbhead.Append("\tTotal");
            sblonghead.Append("\t\t\tTotal");
            sbbase.Append("\t" + basesum);
            memo(sbhead.ToString());
            memo(sbbase.ToString());
            memo("");


            for (int year = baseyear; year <= endyear; year++)
            {
                foreach (string inst in institutiondict.Keys)
                {
                    foreach (programclass pc in programdict.Values)
                    {
                        if (pc.fk)
                        {
                            if (!CBFK.Checked)
                                continue;
                        }
                        else
                        {
                            if (pc.semesters > 2)
                            {
                                if (!CBlongprogram.Checked)
                                    continue;
                            }
                            else
                            {
                                if (!CBshortprogram.Checked)
                                    continue;
                            }
                        }
                        if (pc.yearinstprodrangedict.ContainsKey(year) && (pc.yearinstprodrangedict[year].ContainsKey(inst)))
                            institutiondict[inst].addproductionrange(year, pc.name, pc.yearinstprodrangedict[year][inst]);
                    }

                }
            }


            foreach (string inst in institutiondict.Keys)
            {
                institutiondict[inst].scaleproduction(util.tryconvert(TBxrounds.Text));
            }

            memo(sbhead.ToString());
            for (int year = baseyear; year <= endyear; year++)
            {
                double sum = 0;
                StringBuilder sbline = new StringBuilder(year.ToString());
                foreach (string inst in institutiondict.Keys)
                {
                    sbline.Append("\t" + formatmillions(institutiondict[inst].yearproddict[year].fracmoney));
                    sum += institutiondict[inst].yearproddict[year].fracmoney;
                }
                sbline.Append("\t" + formatmillions(sum));
                memo(sbline.ToString());
            }

            memo("");
            memo("Production range:");
            memo(sbhead.ToString() + sblonghead.ToString());
            for (int year = baseyear; year <= endyear; year++)
            {
                double minsum = 0;
                double maxsum = 0;
                StringBuilder sbline = new StringBuilder(year.ToString());
                foreach (string inst in institutiondict.Keys)
                {
                    Tuple<double, double> tt = institutiondict[inst].yearprodrangedict[year].Range();
                    sbline.Append("\t" + formatmillions(tt.Item1) + " - " + formatmillions(tt.Item2));
                    minsum += tt.Item1;
                    maxsum += tt.Item2;
                }
                sbline.Append("\t" + formatmillions(minsum) + " - " + formatmillions(maxsum));

                sbline.Append("\t");

                double sum = 0;
                foreach (string inst in institutiondict.Keys)
                {
                    Tuple<double, double> tt = institutiondict[inst].yearprodrangedict[year].Range();
                    sbline.Append("\t" + year + "\t" + formatmillions(tt.Item2) + "\t" + formatmillions(tt.Item1) + "\t" + formatmillions(institutiondict[inst].yearproddict[year].fracmoney));
                    sum += institutiondict[inst].yearproddict[year].fracmoney;
                }
                sbline.Append("\t" + year + "\t" + formatmillions(maxsum) + "\t" + formatmillions(minsum) + "\t" + formatmillions(sum));


                memo(sbline.ToString());
            }

            memo("");

            List<string> allareas = new List<string>();

            Dictionary<int, Dictionary<string, double>> progareadict = new Dictionary<int, Dictionary<string, double>>();
            for (int y = baseyear; y <= endyear; y++)
            {
                progareadict.Add(y, new Dictionary<string, double>());
                foreach (programclass pc in programdict.Values)
                {
                    string progarea = pc.area;
                    if (pc.fk)
                    {
                        progarea = "Fristående kurs";
                    }
                    else if (programclass.areanamedict.ContainsKey(progarea))
                        progarea = programclass.areanamedict[progarea];
                    if (!progareadict[y].ContainsKey(progarea))
                    {
                        progareadict[y].Add(progarea, 0);
                        if (!allareas.Contains(progarea))
                            allareas.Add(progarea);
                    }
                    if (pc.yearproddict.ContainsKey(y))
                        progareadict[y][progarea] += pc.yearproddict[y].fracmoney;
                }
            }

            StringBuilder sby = new StringBuilder("");
            for (int y = baseyear; y <= endyear; y++)
                sby.Append("\t" + y);
            memo(sby.ToString());
            foreach (string a in allareas)
            {
                StringBuilder sbl = new StringBuilder(a);
                for (int y = baseyear; y <= endyear; y++)
                {
                    if (progareadict.ContainsKey(y) && progareadict[y].ContainsKey(a))
                        sbl.Append("\t" + formatmillions(progareadict[y][a]));
                    else
                        sbl.Append("\t");
                }
                memo(sbl.ToString());
            }

            memo("");

            //double moneylimit = 1e6;
            double moneylimit = util.tryconvertdouble(TB_moneylimit.Text);

            string othername = "Övriga (<" + formatmillions(moneylimit) + " mnkr)";
            //int y = baseyear+1;
            //int y = baseyear;
            //memo("Produktion " + y);

            foreach (string inst in institutiondict.Keys)
            {
                Dictionary<string, StringBuilder> tabledict = new Dictionary<string, StringBuilder>();
                //memo(inst);
                double sumsmallhst = 0;
                double sumsmallmoney = 0;

                for (int y = baseyear; y <= endyear; y++)
                {
                    var q = from c in institutiondict[inst].progyearproddict[y]
                            orderby c.Value.fracmoney descending
                            select c;
                    foreach (KeyValuePair<string, fracprodclass> c in q)
                    {
                        if (c.Value.fracmoney > moneylimit)
                        {
                            //memo("\t" + c.Key + "\t" + (c.Value.fracmoney / 1e6).ToString("N1") + "\t" + (c.Value.frachst).ToString("N1"));
                            if (!tabledict.ContainsKey(c.Key))
                                tabledict.Add(c.Key, new StringBuilder("\t" + c.Key));
                        }
                        //else
                        //{
                        //    sumsmallhst += c.Value.frachst;
                        //    sumsmallmoney += c.Value.fracmoney;
                        //}
                    }
                }
                //memo("\t"+othername+"\"t" + (sumsmallmoney / 1e6).ToString("N1") + "\t" + (sumsmallhst).ToString("N1"));
                tabledict.Add(othername, new StringBuilder("\t" + othername));

                StringBuilder tableheader = new StringBuilder(inst + "\t");

                for (int yr = baseyear; yr <= endyear; yr++)
                {
                    //memo(inst);
                    tableheader.Append("\t" + yr);
                    var qq = from c in institutiondict[inst].progyearproddict[yr]
                             orderby c.Value.fracmoney descending
                             select c;
                    sumsmallhst = 0;
                    sumsmallmoney = 0;
                    foreach (KeyValuePair<string, fracprodclass> c in qq)
                    {
                        if (tabledict.ContainsKey(c.Key))
                        {
                            tabledict[c.Key].Append("\t" + (c.Value.frachst).ToString("N1"));
                        }
                        else
                        {
                            sumsmallhst += c.Value.frachst;
                            sumsmallmoney += c.Value.fracmoney;
                        }
                    }
                    tabledict[othername].Append("\t" + (sumsmallhst).ToString("N1"));
                }

                memo(tableheader.ToString());
                foreach (string progname in tabledict.Keys)
                {
                    memo(tabledict[progname].ToString());
                }
            }

            //ExamList();

        }

        private void savescenariobutton_Click(object sender, EventArgs e)
        {
            string fn = util.unusedfn(docfolder + @"\scenario-" + DateTime.Now.ToShortDateString() + "-.txt");
            using (StreamWriter sw = new StreamWriter(fn))
            {
                List<string> semlist = make_scenario_semesters(programclass.lastbatchyear, endyear);
                sw.WriteLine(semheader(semlist));
                sw.WriteLine("### Siffrorna anger vilket intag program ska gör varje termin, eller hur många HST fristående kurser.");
                sw.WriteLine("### Ändra siffror för nytt scenario i Excel. Blank ruta = 0 = inget intag. Spara som Unicode-text");
                sw.WriteLine("### Det går bra att lägga till nya program. Fyll bara i lämpliga värden på ny rad.");
                sw.WriteLine("### Ange i institutionskolumnen för nya program vilket gammalt program det liknar mest.");

                var q = programdict.Values.OrderBy(c => c.homeinst);
                foreach (programclass pc in q)
                {
                    StringBuilder sb = new StringBuilder(pc.name + "\t" + pc.semesters + "\t" + pc.averageretention().ToString("N2") + "\t" + pc.homeinst + "\t" + pc.area);
                    foreach (string semsem in semlist)
                    {
                        if (plannedstudents[pc.name].ContainsKey(semsem))
                            sb.Append("\t" + plannedstudents[pc.name][semsem].ToString("N0"));
                        else
                            sb.Append("\t");
                    }
                    sw.WriteLine(sb.ToString());
                }
                memo("Scenario saved to " + fn);
            }
        }

        private void loadscenariobutton_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = docfolder;
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.Title = "Select scenario file";
            Console.WriteLine("opendialog1.Show:");
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fn = openFileDialog1.FileName;
                memo("Reading scenario from " + fn);
                using (StreamReader sr = new StreamReader(fn))
                {
                    string header = sr.ReadLine();
                    string[] hwords = header.Split('\t');
                    int offset = 3;
                    int semcol = 1;
                    int retcol = 2;
                    int instcol = -1;
                    int typecol = 5;

                    for (int i = 1; i < hwords.Length - offset; i++)
                    {
                        if (hwords[i].ToLower().StartsWith("termin"))
                            semcol = i;
                        if (hwords[i].ToLower().StartsWith("reten"))
                            retcol = i;
                        if (hwords[i].ToLower().StartsWith("inst"))
                            instcol = i;
                        if (hwords[i].ToLower().StartsWith("programtyp"))
                            typecol = i;
                        if (hwords[i].ToLower().StartsWith("vt") || hwords[i].ToLower().StartsWith("ht"))
                        {
                            offset = i;
                            break;
                        }
                    }

                    int nline = 0;

                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        nline++;
                        if (nline % 10 == 0)
                            memo("nline = " + nline);

                        if (line.StartsWith("###"))
                            continue;
                        string[] words = line.Split('\t');
                        if (words.Length < offset + 1)
                            continue;

                        string prog = words[0];
                        int sem = util.tryconvert(words[semcol]);
                        double ret = util.tryconvertdouble(words[retcol]);
                        string progarea = "";
                        if (typecol > 0)
                            progarea = words[typecol];
                        Dictionary<string, double> plstuddict = new Dictionary<string, double>();
                        for (int i = offset; i < words.Length; i++)
                        {
                            double stud = util.tryconvertdouble(words[i]);
                            if (stud > 0)
                                plstuddict.Add(hwords[i], stud);
                        }
                        if (prog.Contains("jänst"))
                        {
                            //
                            memo(prog);
                        }
                        if (plannedstudents.ContainsKey(prog))
                            plannedstudents[prog] = plstuddict;
                        else
                            plannedstudents.Add(prog, plstuddict);
                        if (origprogramdict.ContainsKey(prog))
                        {
                            if (sem > origprogramdict[prog].semesters)
                            {
                                origprogramdict[prog].semesters = sem;
                                double rr = origprogramdict[prog].averageretention() > 0 ? origprogramdict[prog].averageretention() : ret;
                                origprogramdict[prog].extendretention(rr);
                            }
                            if (Math.Abs(ret - origprogramdict[prog].averageretention()) > 0.05)
                            {
                                origprogramdict[prog].replaceretention(ret);
                            }
                            if (!string.IsNullOrEmpty(progarea))
                                origprogramdict[prog].area = progarea;
                        }
                        else
                        {
                            string baseprog = words[instcol]; // copy from baseprog in instcol
                            if (origprogramdict.ContainsKey(baseprog))
                            {
                                programclass newprog = origprogramdict[baseprog].clone(false);
                                newprog.name = prog;
                                newprog.semesters = sem;
                                if (!String.IsNullOrEmpty(progarea))
                                    newprog.area = progarea;
                                //newprog.averageaccepted = 0; //No historical admissions
                                if (Math.Abs(ret - newprog.averageretention()) > 0.05)
                                {
                                    newprog.replaceretention(ret);
                                }
                                origprogramdict.Add(prog, newprog);
                            }
                            else
                            {
                                memo("New program " + prog + " with invalid base " + baseprog);
                                origprogramdict.Add(prog, new programclass(prog, sem, ret, progarea));
                                if (!institutiondict.ContainsKey(utaninst))
                                    institutiondict.Add(utaninst, new institutionclass(utaninst));
                            }
                        }
                    }
                    memo(nline + " scenario lines");
                    scenarioloaded = true;
                    businessbutton.Text = "Extrapolate scenario";
                }
            }
        }

        private void examtestbutton_Click(object sender, EventArgs e)
        {
            Dictionary<int, hbookclass> semhistdict = new Dictionary<int, hbookclass>();
            int maxproglength = 11;
            for (int i = 0; i <= maxproglength; i++)
            {
                semhistdict.Add(i, new hbookclass("Exam diff T" + i));
                semhistdict[i].SetBins(-30, 30, 20);
            }
            semhistdict.Add(-1, new hbookclass("Mean exam diff"));
            semhistdict[-1].SetBins(-30, 30, 20);

            foreach (string prog in programdict.Keys)
            {
                if (programdict[prog].fk)
                    continue;

                if (programdict[prog].semesters != 6)
                    continue;

                StringBuilder sb = new StringBuilder(prog);
                for (int sem = 0; sem <= programdict[prog].semesters; sem++)
                {
                    if (programdict[prog].examtransition[sem] != null)
                        sb.Append("\t" + programdict[prog].examtransition[sem].transitionprob.ToString("N3"));
                    else
                        sb.Append("\t");
                }
                memo(sb.ToString());

                foreach (programbatchclass bc in programdict[prog].batchlist)
                {
                    if (bc.actualexam != null) //finns riktiga examensdata?
                    {
                        if (bc.lastrealsemester() >= programdict[prog].semesters) //har kullen hunnit till sista terminen?
                        {
                            int nsem = 0;
                            double examsum = 0;
                            for (int sem = 0; sem <= programdict[prog].semesters; sem++)
                            {
                                double examfc = bc.examforecast(sem, bc.getstud(sem), programdict[prog].examtransition);
                                examsum += examfc;
                                nsem++;
                                semhistdict[sem].Add(examfc - (double)bc.actualexam);
                            }
                            double examfcmean = (examsum / nsem) - (double)bc.actualexam;
                            semhistdict[-1].Add(examfcmean);
                        }
                    }
                }
            }
            for (int i = -1; i <= maxproglength; i++)
            {
                memo(semhistdict[i].GetDHist());
            }
        }

        private void ExamList()
        {
            int baseyear = programclass.lastbatchyear;
            if (endyear < 0)
                endyear = util.tryconvert(TB_endyear.Text);
            //int endyear = baseyear + 5;

            memo("\n======== Examina ===========\n");

            StringBuilder sbyearhead = new StringBuilder();
            for (int yr = baseyear; yr <= endyear; yr++)
                sbyearhead.Append("\t" + yr);
            memo(sbyearhead.ToString() + "\t" + sbyearhead.ToString());

            Dictionary<string, Dictionary<int, double>> examareadict = new Dictionary<string, Dictionary<int, double>>();

            foreach (programclass pc in programdict.Values)
            {
                if (RBteacherexam.Checked)
                    if (pc.area != "L")
                        continue;
                if (RBnursingexam.Checked)
                    if (pc.area != "V")
                        continue;
                if (pc.examsum(baseyear, endyear) == 0)
                    continue;

                StringBuilder sb = new StringBuilder(pc.name);

                if (!examareadict.ContainsKey(pc.area))
                    examareadict.Add(pc.area, new Dictionary<int, double>());
                //Average # exams
                for (int yr = baseyear; yr <= endyear; yr++)
                {
                    if (pc.examforecastrangedict.ContainsKey(yr))
                    {
                        sb.Append("\t" + pc.examforecastrangedict[yr].Average());
                        if (!examareadict[pc.area].ContainsKey(yr))
                            examareadict[pc.area].Add(yr, pc.examforecastrangedict[yr].Average());
                        else
                            examareadict[pc.area][yr] += pc.examforecastrangedict[yr].Average();
                    }
                    else
                        sb.Append("\t");
                }

                //Range:
                sb.Append("\t" + pc.name);
                for (int yr = baseyear; yr <= endyear; yr++)
                {
                    if (pc.examforecastrangedict.ContainsKey(yr))
                        sb.Append("\t'" + pc.examforecastrangedict[yr].RangeString());
                    else
                        sb.Append("\t");
                }
                memo(sb.ToString());
            }

            foreach (string ar in examareadict.Keys)
            {
                string arstring = "Övriga";
                if (ar == "L")
                    arstring = "Alla lärarutbildningar";
                else if (ar == "V")
                    arstring = "Alla vårdutbildningar";
                StringBuilder sb = new StringBuilder(arstring);
                for (int yr = baseyear; yr <= endyear; yr++)
                {
                    if (examareadict[ar].ContainsKey(yr))
                    {
                        sb.Append("\t" + examareadict[ar][yr]);
                    }
                    else
                        sb.Append("\t");

                }
                memo(sb.ToString());
            }
        }

        private void buttonExamforecast_Click(object sender, EventArgs e)
        {
            ExamList();

        }

        private void TBxrounds_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        int findlimitindex(double value, int[] limit)
        {
            for (int i = 0; i < limit.Length - 1; i++)
            {
                if (value < limit[i + 1])
                    return i;
            }
            return limit.Length - 1;
        }

        private void AppRegButton_Click(object sender, EventArgs e)
        {
            int regmin = 20;
            memo("======= Sökande till registrerade ==============");
            memo("Hur många 1:ahandssökande krävs för att få " + regmin + " registrerade nybörjare?");
            foreach (programclass pc in origprogramdict.Values)
            {
                if (pc.appltransition[1] != null)
                {
                    double apptrans = pc.appltransition[1].transitionprob;
                    if (apptrans > 0)
                    {
                        memo(pc.name + "\t" + Math.Round(regmin / apptrans));
                    }
                }
            }

            bool lastsem = true;
            sectorhist("", lastsem, false);
            foreach (string sector in sectorlist)
                sectorhist(sector, lastsem, false);
            sectorhist("", lastsem, true);
            foreach (string sector in sectorlist)
                sectorhist(sector, lastsem, true);

            lastsem = false;
            sectorhist("", lastsem, false);
            foreach (string sector in sectorlist)
                sectorhist(sector, lastsem, false);

            int minstud = 20;
            memo("\tInst\tMedel-nybörjare\tRisk <" + minstud + " nybörjare\tMedel alla terminer\tRisk <" + minstud + " alla terminer");
            foreach (programclass pc in origprogramdict.Values)
            {
                if (pc.appltransition[1] != null)
                {
                    StringBuilder sb = new StringBuilder(pc.name + "\t" + pc.homeinst);
                    int nbmean = 0;
                    int nbmeanbad = 0;
                    double? nbmeansum = 0;
                    int nb1 = 0;
                    int nb1bad = 0;
                    double nb1sum = 0;
                    foreach (programbatchclass pb in pc.batchlist)
                    {
                        nb1++;
                        nb1sum += pb.getstud(1);
                        if (pb.getstud(1) < minstud)
                            nb1bad++;
                        if (pb.getactualstud(pc.semesters) != null)
                        {
                            nbmean++;
                            nbmeansum += pb.getmeanstud(1, pc.semesters);
                            if (pb.getmeanstud(1, pc.semesters) < minstud)
                                nbmeanbad++;
                        }
                    }
                    double b1 = nb1sum / nb1;
                    double? bmean = (nbmean > 0) ? nbmeansum / nbmean : null;
                    double bad1frac = 100*nb1bad / (double)nb1;
                    double? badmeanfrac = (nbmean > 0) ? 100*(nbmeanbad / (double?)nbmean) : null;
                    sb.Append("\t" + b1 + "\t" + bad1frac + "%\t" + (bmean != null ? bmean.ToString() : "") + "\t" + (badmeanfrac != null ? badmeanfrac.ToString() : "")+"%");
                    memo(sb.ToString());
                }
            }

        }

        private string limitheader(int i, int[] limit)
        {
            if (i == 0)
                return "<" + limit[1];
            else if (i == limit.Length - 1)
                return limit.Last() + "+";
            else
                return "'" + limit[i] + "-" + (limit[i + 1] - 1);

        }

        private void sectorhist(string sector, bool lastsem, bool meanstud)
        {
            int[] applimit = new int[] { -1, 10, 15, 20, 25, 30, 35, 40, 45, 50, 60, 80, 100 };
            int[] reglimit = new int[] { -1, 10, 15, 20, 25, 30, 35, 40, 45, 50, 60, 80, 100 };

            List<string> progskip = new List<string>() { "Vidareutbildning av lärare" };

            int[,] hist = new int[applimit.Length, reglimit.Length];

            memo("");
            string s = lastsem ? (meanstud ? "stud i genomsnitt alla terminer" : "stud sista terminen") : "nybörjare(+reserver)";
            memo("======= Histogram sökande vs " + s + " ==============");
            memo("Hur många 1:ahandssökande krävs för att få ett visst antal registrerade " + s + "?");
            if (String.IsNullOrEmpty(sector))
                memo("======= alla utbildningar ==========");
            else
                memo("======= sektor " + sector + " ============");

            foreach (programclass pc in origprogramdict.Values)
            {
                if (!String.IsNullOrEmpty(sector) && pc.sector != sector)
                    continue;
                if (progskip.Contains(pc.name))
                    continue;
                if (pc.appltransition[1] != null)
                {
                    foreach (programbatchclass pb in pc.batchlist)
                    {
                        int iapp = findlimitindex(pb.applicants, applimit);
                        double? stud = lastsem ? pb.getactualstud(pc.semesters) : pb.getstud(0) + pb.reserves;
                        if (stud == null)
                            continue;
                        if (lastsem && meanstud)
                        {
                            stud = pb.getmeanstud(1, pc.semesters);
                        }
                        int ireg = findlimitindex((double)stud, reglimit);
                        hist[iapp, ireg]++;
                        if (lastsem && iapp == 0 && ireg > 6)
                        {
                            memo(pc.name + "\t" + pb.batchstart + "\t" + pb.applicants + "\t" + stud);
                        }
                    }
                }
            }

            StringBuilder sbhead = new StringBuilder("\t" + limitheader(0, reglimit));
            for (int i = 1; i < reglimit.Length; i++)
                sbhead.Append("\t" + limitheader(i, reglimit));
            memo(sbhead.ToString());

            for (int i = 0; i < applimit.Length; i++)
            {
                StringBuilder sb = new StringBuilder(limitheader(i, applimit));
                for (int j = 0; j < reglimit.Length; j++)
                {
                    sb.Append("\t" + hist[i, j]);
                }
                memo(sb.ToString());
            }
        }

        private List<string> parsecoursecodes(string input)
        {
            List<string> ls = new List<string>();

            string rex = @"\w\w[\w\d]{4,5}";

            foreach (Match m in Regex.Matches(input,rex))
            {
                if (m.Value.ToUpper() == m.Value)
                    ls.Add(m.Value);
            }
            return ls;
        }

        private string getsubjectcode(string code)
        {
            
            string rex1 = @"[GAB]([\p{Lu}][\p{Lu}])\w+";
            string rex2 = @"([\p{Lu}][\p{Lu}])\w+";
            foreach (Match m in Regex.Matches(code,rex1))
            {
                return m.Groups[1].Value;
            }
            foreach (Match m in Regex.Matches(code, rex2))
            {
                return m.Groups[1].Value;
            }
            return "";
        }

        programclass findprogram(string code)
        {
            string c2 = code;
            if (programkoppling.ContainsKey(code))
                c2 = programkoppling[code];
            return findcourse(c2, origprogramdict,new Dictionary<string, programclass>());
        }

        programclass findcourse(string code)
        {
            return findcourse(code, fkdict,fkcodedict);
        }

        programclass findcourse(string code, Dictionary<string,programclass> cdict, Dictionary<string, programclass> codedict) //either code or name as input
        {
            if (codedict.ContainsKey(code))
                return codedict[code];
            if (code.Length == 6)
            {
                var q = from c in cdict.Values
                        where c.coursecodelist.Contains(code)
                        select c;
                var cc = q.FirstOrDefault();
                if (cc != null)
                    return cc;
                else if (cdict.ContainsKey(code))
                    return cdict[code];
                else
                    return null;
            }
            else if (code.Length == 5)
            {
                var q = from c in cdict.Values
                        where c.applcodelist.Contains(code)
                        select c;
                var cc = q.FirstOrDefault();
                if (cc != null)
                    return cc;
                else if (cdict.ContainsKey(code))
                    return cdict[code];
                else
                    return null;
            }
            else if (cdict.ContainsKey(code))
                return cdict[code];
            else
            {
                var q = from c in cdict.Values
                        where c.coursecodelist.Contains(code)
                        select c;
                return q.FirstOrDefault();
            }
        }

        private void read_aktiva_utb_file()
        {
            hbookclass typehist = new hbookclass("Utbildningstyp");
            openFileDialog1.InitialDirectory = folder;
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.Title = "Select aktiva_utb_tabell file";
            Console.WriteLine("opendialog1.Show:");
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fn = openFileDialog1.FileName;
                memo("Reading aktiva utb from " + fn);
                using (StreamReader sr = new StreamReader(fn))
                {
                    sr.ReadLine(); //throw away header line
                    sr.ReadLine(); //throw away header line
                    sr.ReadLine(); //throw away header line

                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        string[] words = line.Split('\t');
                        if (String.IsNullOrEmpty(words[0].Trim()))
                            continue;
                        double hp = util.tryconvertdouble(words[3]);
                        string name = words[2];
                        string applcode = words[0];
                        string coursecode = words[1];
                        string utype = words[8];
                        typehist.Add(utype);
                        double fullfee = util.tryconvertdouble(words[16].Replace(" ", ""));
                        if (utype == "Kurs")
                        {
                            programclass fk = findcourse(name);
                            if (fk == null && !String.IsNullOrEmpty(applcode))
                                fk = findcourse(applcode);
                            if (fk == null && !String.IsNullOrEmpty(coursecode))
                                fk = findcourse(coursecode);
                            if (fk == null) //create new entry
                            {
                                fk = new programclass(name);
                                fk.name = name;
                                fkdict.Add(name, fk);
                            }
                            if (fk.hp <= 0)
                                fk.hp = hp;
                            fk.semesters = 1;
                            fk.utype = utype;
                            fk.fee = fullfee;
                            fk.fk = true;
                            fk.subjectcode = getsubjectcode(coursecode);
                            if (subjinstdict.ContainsKey(fk.subjectcode))
                                fk.homeinst = shortinstdict[subjinstdict[fk.subjectcode]];
                            if (!String.IsNullOrEmpty(words[10]))
                                fk.partofpackage = words[10].Trim(',');
                            if (!fkcodedict.ContainsKey(coursecode))
                                fkcodedict.Add(coursecode, fk);
                            if (!fkcodedict.ContainsKey(applcode))
                                fkcodedict.Add(applcode, fk);
                            if (!fk.coursecodelist.Contains(coursecode))
                            {
                                fk.coursecodelist.Add(coursecode);
                                fk.subjectcode = getsubjectcode(coursecode);
                                fk.homeinst = (shortinstdict[subjinstdict[fk.subjectcode]]);
                            }
                            if (!fk.applcodelist.Contains(applcode))
                            {
                                fk.applcodelist.Add(applcode);
                            }
                        }
                        else
                        {
                            programclass pc = findprogram(name);
                            if (pc == null && !String.IsNullOrEmpty(applcode))
                                pc = findprogram(applcode);
                            if (pc == null) //create new entry
                            {
                                pc = new programclass(name);
                                pc.name = name;
                                origprogramdict.Add(name, pc);
                            }
                            pc.hp = hp;
                            pc.semesters = (int)Math.Ceiling(hp / 30);
                            pc.utype = utype;
                            pc.fee = fullfee;
                            pc.fk = false;
                            if (!String.IsNullOrEmpty(applcode) && !pc.applcodelist.Contains(applcode))
                            {
                                pc.applcodelist.Add(applcode);
                            }
                        }
                    }
                }
                memo("# courses = " + fkdict.Count);
            }

            memo(typehist.GetSHist());

            foreach (string s in fkdict.Keys)
            {
                fkdict[s].calculate_transitions();
            }

            var qpart = from c in fkdict.Values
                        where !String.IsNullOrEmpty(c.partofpackage)
                        select c;
            foreach (programclass part in qpart)
            {
                programclass paket = findprogram(part.partofpackage);
                if (paket == null)
                {
                    memo(part.partofpackage + " not found");
                }
                else
                {
                    paket.homeinst = part.homeinst;
                    if (!paket.coursedict.ContainsKey(1))
                        paket.coursedict.Add(1, new Dictionary<string, double>());
                    paket.coursedict[1].Add(part.bestcode(), 1);
                }
            }

        }


        string hprex = @" \d+(\.\d)? hp";

        private void read_fkfile()
        {
            openFileDialog1.InitialDirectory = folder;
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.Title = "Select FK file";
            Console.WriteLine("opendialog1.Show:");
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fn = openFileDialog1.FileName;
                memo("Reading FK from " + fn);
                using (StreamReader sr = new StreamReader(fn))
                {
                    sr.ReadLine(); //throw away header line

                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        string[] words = line.Split('\t');
                        double hp = -1;
                        string name = words[4];
                        string code = words[3];
                        foreach (Match m in Regex.Matches(words[4],hprex))
                        {
                            hp = util.tryconvertdouble(m.Value.Trim().Replace(" hp",""));
                            name = name.Replace(m.Value, "").Trim();
                        }
                        if (hp < 0)
                        {
                            memo("HP fail " + name);
                        }
                        programclass fk = findcourse(code);
                        if (fk == null)
                            fk = findcourse(name);
                        if (fk == null) //create new entry
                        {
                            fk = new programclass(name);
                            fk.name = name;
                            fk.semesters = 1;
                            if (fk.hp <= 0 && hp > 0)
                                fk.hp = hp;
                            fk.subject = words[0];
                            fk.sector = words[1];
                            fkdict.Add(name, fk);
                        }
                        //public programbatchclass(double?[] actualstud, int prog, string bstart, int exam, int appl, int res) //real data
                        double?[] actualstud = new double?[2];
                        actualstud[0] = util.tryconvert(words[6]);
                        actualstud[1] = util.tryconvert(words[7]);
                        programbatchclass kt = new programbatchclass(actualstud, -1, util.semester3to2(words[2]), util.tryconvert(words[8]), util.tryconvert(words[5]), 0);
                        fk.batchlist.Add(kt);
                        if (!fk.coursecodelist.Contains(words[3]))
                        {
                            fk.coursecodelist.Add(words[3]);
                            fk.subjectcode = getsubjectcode(words[3]);
                            fk.homeinst = (shortinstdict[subjinstdict[fk.subjectcode]]);
                        }
                        if (!fkcodedict.ContainsKey(words[3]))
                            fkcodedict.Add(words[3], fk);

                    }
                }
                memo("# courses = " + fkdict.Count);
            }



        }

        private void FKbutton_Click(object sender, EventArgs e)
        {
            read_fkfile();

            Dictionary<int, courseroomclass> roomdict = new Dictionary<int, courseroomclass>();
            Dictionary<string, List<int>> coderoomdict = new Dictionary<string, List<int>>();

            openFileDialog1.Title = "Select Learn file";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            Console.WriteLine("opendialog1.Show:");
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fn = openFileDialog1.FileName;
                memo("Reading Learn rooms from " + fn);
                using (StreamReader sr = new StreamReader(fn))
                {
                    sr.ReadLine(); //throw away header line
                    int nroom = 0;
                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        string[] words = line.Split('\t');

                        //memo("-----"+words[0]);
                        courseroomclass crc = new courseroomclass();
                        nroom++;
                        foreach (string s in parsecoursecodes(words[0]))
                        {
                            //memo(s);
                            crc.courses.Add(s);
                            programclass pc = findcourse(s);
                            if (pc == null)
                                crc.nodata++;
                            else
                            {
                                programbatchclass pbc = pc.getbatch("HT22");
                                if (pbc != null)
                                crc.knownstudents += (int)pbc.getstud(1);
                            }
                            if (!coderoomdict.ContainsKey(s))
                                coderoomdict.Add(s, new List<int>());
                            coderoomdict[s].Add(nroom);
                        }
                        roomdict.Add(nroom, crc);
                    }
                }
                memo("# courses = " + fkdict.Count);
            }

            int minstud = 20;

            memo("\tÄmneskod\tSektor\tHT22\tVT22\tHT21\tVT21\tHT20\tVT20\tHT19\tVT19\tMedel-ffgreg\tRisk för <"+minstud+" stud\tnbadaftergood\t<"+minstud+" stud\t>="+minstud+" stud");
            foreach (string name in fkdict.Keys)
            {
                //if (fkdict[name].batchlist.Count < 3)
                //    continue;
                int nshareroom = 0;
                foreach (string code in fkdict[name].coursecodelist)
                {
                    if (coderoomdict.ContainsKey(code))
                    {
                        foreach (int nroom in coderoomdict[code])
                        nshareroom += (roomdict[nroom].courses.Count-1);
                    }
                }
                int ngood = 0;
                int nbad = 0;
                int nbadaftergood = 0;
                programbatchclass pb = fkdict[name].getfirstbatch();
                while (pb != null)
                {
                    if (pb.getstud(1) < minstud)
                    {
                        nbad++;
                        nbadaftergood++;
                    }
                    else
                    {
                        ngood++;
                        nbadaftergood = 0;
                    }
                    pb = fkdict[name].getnextbatch(pb.batchstart);
                }
                StringBuilder sb = new StringBuilder(name+"\t"+fkdict[name].subjectcode+"\t"+subjinstdict[fkdict[name].subjectcode]);
                string startsem = "HT22";
                int nsem = 8;
                double regsum = 0;
                double regcount = 0;
                for (int i=0;i<nsem;i++)
                {
                    programbatchclass pb2 = fkdict[name].getbatch(startsem);
                    if (pb2 != null)
                    {
                        int nstud = (int)pb2.getstud(1);
                        regsum += nstud;
                        regcount++;
                        sb.Append("\t" + nstud);
                    }
                    else
                        sb.Append("\t");
                    startsem = util.decrementsemester(startsem);
                }
                string meanreg = regcount > 0 ? (regsum / regcount).ToString("N1") : "";
                sb.Append("\t"+meanreg);
                sb.Append("\t" + ((100*nbad)/(double)(nbad+ngood)).ToString("N0") + "%\t" + nbadaftergood + "\t" + nbad + "\t" + ngood);
                memo(sb.ToString());
            }

        }

        private List<dictclass> read_hst_hpr(string fn)
        {
            List<dictclass> courses = new List<dictclass>();

            memo("Reading hst/hpr data from " + fn);

            using (StreamReader sr = new StreamReader(fn))
            {
                sr.ReadLine(); //throw away header line
                sr.ReadLine(); //throw away header line
                sr.ReadLine(); //throw away header line
                string header = sr.ReadLine();
                string[] hwords = header.Split('\t');
                int nline = 0;
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    string[] words = line.Split('\t');
                    nline++;
                    dictclass d = new dictclass(hwords, words);

                    if (!d.Has("Kurskod"))
                        continue;
                    if (!d.Has("HP"))
                        continue;
                    courses.Add(d);
                }
            }
            memo("# courses = " + courses.Count);

            return courses;
        }

        private void HSTbutton_Click(object sender, EventArgs e)
        {
            List<string> subjectcodes = new List<string>();
            List<dictclass> courses = new List<dictclass> ();

            openFileDialog1.Title = "Select hst_hpr_utfall_budget file";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            Console.WriteLine("opendialog1.Show:");
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fn = openFileDialog1.FileName;
                memo("Reading hst/hpr data from " + fn);
                
                using (StreamReader sr = new StreamReader(fn))
                {
                    sr.ReadLine(); //throw away header line
                    sr.ReadLine(); //throw away header line
                    sr.ReadLine(); //throw away header line
                    string header = sr.ReadLine();
                    string[] hwords = header.Split('\t');
                    int nline = 0;
                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        string[] words = line.Split('\t');
                        nline++;
                        dictclass d = new dictclass(hwords, words);

                        if (!d.Has("Kurskod"))
                            continue;
                        string subjectcode = getsubjectcode(d.Get("Kurskod"));
                        if (String.IsNullOrEmpty(subjectcode))
                            continue;
                        if (!subjectcodes.Contains(subjectcode))
                            subjectcodes.Add(subjectcode);
                        courses.Add(d);
                    }
                }
                memo("# courses = " + courses.Count);
            }

            string utfallstring = "HST utfall";
            string budgetstring = "Budget HST";
            var q00 = from c in courses
                           where c.Getdouble(budgetstring) == 0
                           where c.Getdouble(utfallstring) == 0
                           select c;
            var q0budget = from c in courses 
                           where c.Getdouble(budgetstring) == 0
                           where c.Getdouble(utfallstring) > 0
                           select c;
            var q0utfall = from c in courses
                           where c.Getdouble(budgetstring) > 0
                           where c.Getdouble(utfallstring) == 0 
                           select c;
            var qnonzero = from c in courses 
                           where c.Getdouble(budgetstring) > 0
                           where c.Getdouble(utfallstring) > 0
                           select c;

            List<string> toprint = new List<string>() { "Kurskod", "HST utfall", "Budget HST", "Kurs" };
            StringBuilder sbhead = new StringBuilder();
            foreach (string s in toprint)
            {
                sbhead.Append(s + "\t");
            }
            memo("\n== Kurser med noll i budget och noll i utfall: ==");
            memo(sbhead.ToString());
            foreach (dictclass d in q0budget)
            {
                memo(d.Printline(toprint));
            }

            memo("\n== Oplanerade kurser med noll i budget: ==");
            memo(sbhead.ToString());
            foreach (dictclass d in q0budget)
            {
                memo(d.Printline(toprint));
            }

            memo("\n== Inställda kurser med noll i utfall: ==");
            memo(sbhead.ToString());
            foreach (dictclass d in q0utfall)
            {
                memo(d.Printline(toprint));
            }

            int nbetter = 0;
            memo("\n== Kurser med mycket HÖGRE utfall än budget: ==");
            memo(sbhead.ToString());
            foreach (dictclass d in qnonzero)
            {
                double utfall = d.Getdouble("HST utfall");
                double budget = d.Getdouble("Budget HST");
                if (utfall - budget > 5 || utfall / budget > 2)
                {
                    memo(d.Printline(toprint));
                    nbetter++;
                }
            }

            int nworse = 0;
            memo("\n== Kurser med mycket LÄGRE utfall än budget: ==");
            memo(sbhead.ToString());
            foreach (dictclass d in qnonzero)
            {
                double utfall = d.Getdouble("HST utfall");
                double budget = d.Getdouble("Budget HST");
                if (utfall - budget < -5 || utfall / budget < 0.5)
                {
                    memo(d.Printline(toprint));
                    nworse++;
                }
            }

            memo("");

            memo("Oplanerade kurser:\t" + q0budget.Count());
            memo("Inställda kurser:\t" + q0utfall.Count());
            memo("Kurser med mycket HÖGRE utfall än budget:\t" + nbetter);
            memo("Kurser med mycket LÄGRE utfall än budget:\t" + nworse);
            memo("Kurser genomförda enligt budget:\t" + (courses.Count - nbetter - nworse - q0budget.Count() - q0utfall.Count()));
            memo("Totalt antal kurser:\t" + courses.Count);
        }

        private void Excelplanningbutton_Click(object sender, EventArgs e)
        {
            ExcelForm ef = new ExcelForm();
            ef.Show();
        }

        private void read_program_programkurser_intag()
        {
            List<string> progs = new List<string>();
            openFileDialog1.Title = "Select Program_programkurser_intag file";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fn = openFileDialog1.FileName;
                memo("Reading program/course data from " + fn);

                int nfoundcode = 0;
                int nfoundname = 0;
                int nnfound = 0;
                int nprognotfound = 0;
                using (StreamReader sr = new StreamReader(fn))
                {
                    string header = sr.ReadLine();
                    string[] hwords = header.Split('\t');
                    sr.ReadLine(); //throw away header line
                    int nline = 0;
                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        string[] words = line.Split('\t');
                        nline++;
                        string prog = words[0];
                        string code = words[1];
                        if (string.IsNullOrEmpty(code))
                            continue;
                        string coursename = words[2];
                        int sem = util.tryconvert(words[3]);
                        int stud = util.tryconvert(words[4]);
                        if (!progs.Contains(prog))
                            progs.Add(prog);
                        programclass pc = findprogram(prog);
                        if (pc == null)
                        {
                            memo(prog + " not found");
                            nprognotfound++;
                            continue;
                        }
                        programclass course = findcourse(code);
                        if (course == null)
                        {
                            course = findcourse(coursename);
                            //memo(coursename + " sought by name");
                            if (course != null)
                            {
                                nfoundname++;
                                course.coursecodelist.Add(code);
                                fkcodedict.Add(code, course);
                            }
                        }
                        else
                            nfoundcode++;
                        if (course == null)
                        {
                            //memo(coursename + " not found");
                            nnfound++;
                            course = new programclass(coursename);
                            course.name = coursename;
                            course.semesters = 1;
                            //course.hp = hp;
                            //course.subject = words[0];
                            //course.sector = words[1];
                            course.coursecodelist.Add(code);
                            course.subjectcode = getsubjectcode(code);
                            course.homeinst = (shortinstdict[subjinstdict[course.subjectcode]]);
                            fkdict.Add(coursename, course);
                            fkcodedict.Add(code, course);
                            //continue;
                        }
                        string bestcode = course.bestcode();
                        if (!pc.coursedict.ContainsKey(sem))
                            pc.coursedict.Add(sem, new Dictionary<string, double>());
                        if (!pc.coursedict[sem].ContainsKey(bestcode))
                            pc.coursedict[sem].Add(bestcode, stud);
                        else
                            pc.coursedict[sem][bestcode] += stud;
                    }
                    memo("# lines " + nline);
                    memo("# progs = " + progs.Count);
                    memo("#courses found by code: " + nfoundcode);
                    memo("#courses found by name: " + nfoundname);
                    memo("#courses not found: " + nnfound);
                    memo("#programs not found: " + nprognotfound);

                }

                //double hpnormsum = 27;
                //foreach (string prog in origprogramdict.Keys)
                //{
                //    programclass pc = origprogramdict[prog];
                //    double[] tstud = pc.batchsemsum(beforebatch);
                //    foreach (int sem in pc.coursedict.Keys.ToList())//=1;sem<=pc.semesters;sem++)
                //    {
                //        double hpsum = 0;
                //        if (!pc.coursedict.ContainsKey(sem))
                //            continue;
                //        foreach (string code in pc.coursedict[sem].Keys.ToList())
                //        {
                //            double normstud;
                //            if (pc.utype == "Kurspaket")
                //                normstud = 0.9;
                //            else
                //            {
                //                normstud = pc.coursedict[sem][code] / tstud[sem];
                //                if (normstud > 1)
                //                    normstud = 1;
                //            }
                //            //if (pc.coursedict[sem][code] > tstud[sem])
                //            //{
                //            //    memo(prog + "\t" + code + "\t" + sem + "\t" + pc.coursedict[sem][code] + "\t" + tstud[sem]);
                //            //}
                //            if (normstud < 0.01 || pc.coursedict[sem][code] <= 1)
                //            {
                //                zeroed++;
                //                normstud = 0;
                //            }
                //            else
                //            {
                //                nonzeroed++;
                //                fkcodedict[code].activecourse = true;
                //                if (fkcodedict[code].hp > 0)
                //                    hpsum += normstud * fkcodedict[code].hp;
                //                else
                //                    memo(code + "\t" + fkcodedict[code].name);
                //            }
                //            pc.coursedict[sem][code] = normstud;
                //        }
                //        //memo(prog + "\t" + sem + "\t" + hpsum.ToString("N1"));
                //        foreach (string code in pc.coursedict[sem].Keys.ToList())
                //        {
                //            if (pc.utype == "Kurspaket")
                //                continue;
                //            if (pc.coursedict[sem][code] == 0)
                //                continue;
                //            pc.coursedict[sem][code] *= hpnormsum/hpsum;
                //        }
                //    }
                //}

            }
        }

        private void normalize_pccoursedict()
        {
            memo("Normalizing:");
            string beforebatch = "VT16";
            int zeroed = 0;
            int nonzeroed = 0;
            double hpnormsum = 27;
            foreach (string prog in origprogramdict.Keys)
            {
                programclass pc = origprogramdict[prog];
                double[] tstud = pc.batchsemsum(beforebatch);
                foreach (int sem in pc.coursedict.Keys.ToList())//=1;sem<=pc.semesters;sem++)
                {
                    double hpsum = 0;
                    if (!pc.coursedict.ContainsKey(sem))
                        continue;
                    foreach (string code in pc.coursedict[sem].Keys.ToList())
                    {
                        double normstud;
                        if (pc.utype == "Kurspaket")
                            normstud = 0.9;
                        else
                        {
                            normstud = pc.coursedict[sem][code] / tstud[sem];
                            if (normstud > 1)
                                normstud = 1;
                        }
                        //if (pc.coursedict[sem][code] > tstud[sem])
                        //{
                        //    memo(prog + "\t" + code + "\t" + sem + "\t" + pc.coursedict[sem][code] + "\t" + tstud[sem]);
                        //}
                        if ((normstud < 0.01 || pc.coursedict[sem][code] <= 1) && pc.utype != "Kurspaket")
                        {
                            zeroed++;
                            normstud = 0;
                        }
                        else
                        {
                            nonzeroed++;
                            fkcodedict[code].activecourse = true;
                            if (fkcodedict[code].hp > 0)
                                hpsum += normstud * fkcodedict[code].hp;
                            else
                                memo(code + "\t" + fkcodedict[code].name);
                        }
                        pc.coursedict[sem][code] = normstud;
                    }
                    //memo(prog + "\t" + sem + "\t" + hpsum.ToString("N1"));
                    foreach (string code in pc.coursedict[sem].Keys.ToList())
                    {
                        if (pc.utype == "Kurspaket")
                            continue;
                        if (pc.coursedict[sem][code] == 0)
                            continue;
                        pc.coursedict[sem][code] *= hpnormsum / hpsum;
                    }
                }
            }
            memo("# nonzeroed " + nonzeroed);
            memo("# zeroed " + zeroed);

        }

        private Dictionary<string,double> parse_utbomr(string s)
        {
            //LU 50, HU 50
            Dictionary<string, double> dict = new Dictionary<string, double>();
            string[] ss = s.Split(',');
            foreach (string sss in ss)
            {
                string[] s4 = sss.Trim().Split();
                if (s4.Length < 2)
                    continue;
                if (s4[0] == "VÃ…")
                    s4[0] = "VÅ";
                if (!lokal_ers_hpr.ContainsKey(s4[0]))
                    memo(s);
                double frac = 0.01 * util.tryconvertdouble(s4[1]);
                if (s4[0] == "MM") //Media 352 hst egentligen, bara 20 tillåtna
                {
                    double frac2 = (frac * (352 - 20)) / 352;
                    frac = (frac * 20) / 352;
                    dict.Add("TE", frac2);
                }
                dict.Add(s4[0], frac);

                
            }
            return dict;
        }

        public static double hstkr(double hst, Dictionary<string,double> hstpeng)
        {
            double kr = 0;
            foreach (string area in hstpeng.Keys)
            {
                kr += lokal_ers_hst[area] * hstpeng[area];
            }
            return kr * hst;
        }

        public static double hprkr(double hst, Dictionary<string, double> hstpeng)
        {
            double kr = 0;
            foreach (string area in hstpeng.Keys)
            {
                kr += lokal_ers_hpr[area] * hstpeng[area];
            }
            return kr * hst;
        }

        private void do_hst_hpr_dict(List<dictclass> dict)
        {
            foreach (dictclass d in dict)
            {
                string code = d.Get("Kurskod");
                if (code == "MT1051")
                    code += "";
                string name = d.Get("Kurs");
                programclass course = findcourse(code);
                if (course == null)
                    course = findcourse(name);
                if (course == null)
                    continue;
                if (course.hp <= 0)
                    course.hp = util.tryconvertdouble(d.Get("HP"));
                if (course.studentpengarea.Count == 0 && d.Has("UtbOmr"))
                    course.studentpengarea = parse_utbomr(d.Get("UtbOmr"));
                if (!d.Has("HST utfall"))
                    continue;
                double hst = util.tryconvertdouble(d.Get("HST utfall"));
                if (hst > 0)
                {
                    double hpr = util.tryconvertdouble(d.Get("HPR utfall"));
                    double krhst = hstkr(hst, course.studentpengarea);
                    double krhpr = hprkr(hpr, course.studentpengarea);
                    double kr = krhst + krhpr;
                    course.totalprod.add(hst, hpr, krhst, krhpr, kr);
                    course.activecourse = true;
                }
            }

            foreach (programclass course in fkdict.Values)
            {
                if (course.studentpengarea.Count == 0) //ta studentpeng från annan kurs i samma ämne
                {
                    programclass c2 = (from c in fkdict.Values
                                      where c.subjectcode == course.subjectcode
                                      select c).First();
                    course.studentpengarea = c2.studentpengarea;
                }
            }

        }

        private void do_hst_hpr()
        {
            for (int i=2019;i<2023;i++)
            {
                string fn = folder + @"\hst_hpr_utfall_budget_reg " + i + ".txt";
                var dict = read_hst_hpr(fn);
                do_hst_hpr_dict(dict);
            }
            string fnmiss = folder + @"\missing-courses.txt";
            var dmiss = read_hst_hpr(fnmiss);
        }

        private void coursedatabutton_Click(object sender, EventArgs e)
        {
            read_aktiva_utb_file();
            read_fkfile();
            read_program_programkurser_intag();
            do_hst_hpr();
            normalize_pccoursedict();
        }

        private double[] batchsemsum(string beforebatch, programclass pc)
        {
            return pc.batchsemsum(beforebatch);
        }

        private void testbutton_Click(object sender, EventArgs e)
        {
            //List courses without HP
            foreach (var c in fkdict.Values)
            {
                if (c.hp <= 0 && c.activecourse)
                    memo(c.bestcode() + "\t" + c.name);
            }
                


            //test batchsemsum
            //string beforebatch = "VT16";
            //foreach (string prog in origprogramdict.Keys)
            //{
            //    double[] tstud = batchsemsum(beforebatch, origprogramdict[prog]);
            //    StringBuilder sb = new StringBuilder(prog);
            //    for (int i = 1; i <= origprogramdict[prog].semesters; i++)
            //        sb.Append("\t" + tstud[i]);
            //    memo(sb.ToString());
            //}
        }
    }
}
