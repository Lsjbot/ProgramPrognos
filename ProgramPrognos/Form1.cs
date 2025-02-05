﻿using System;
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
using System.Globalization;

namespace ProgramPrognos
{
    public partial class Form1 : Form
    {
        //public static string homefolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        public static string homefolder = Environment.GetEnvironmentVariable("onedrive")+@"\Dokument";
        public static string folder = homefolder + @"\Invärld\Långtidsbudget";
        public static string docfolder = homefolder;//@"\\dustaff\home\"+Environment.UserName+@"\Documents";
        public static string utaninst = "Utan institution";
        public static string hda = "HDa";

        public static List<string> sectorlist = new List<string>();


        public static Dictionary<string, string> programkoppling = new Dictionary<string, string>();
        public static Dictionary<string, programclass> programdict = new Dictionary<string, programclass>();
        public static Dictionary<string, programclass> origprogramdict = new Dictionary<string, programclass>();
        public static Dictionary<string, institutionclass> institutiondict = new Dictionary<string, institutionclass>();
        public static Dictionary<string, string> instshortdict = new Dictionary<string, string>();
        public static Dictionary<string, string> shortinstdict = new Dictionary<string, string>();
        public static Dictionary<string, string> subjinstdict = new Dictionary<string, string>(); // from coursecode subjects to institutions
        public static Dictionary<string, string> subjnamedict = new Dictionary<string, string>(); // from coursecode subjects to subject name
        public static bool scenarioloaded = false;

        public static Dictionary<string,programclass> fkdict = new Dictionary<string,programclass>();
        public static Dictionary<string, programclass> fkcodedict = new Dictionary<string, programclass>();
        public static Dictionary<string, programclass> paketdict = new Dictionary<string, programclass>();

        public static Dictionary<string, double> lokal_ers_hst = new Dictionary<string, double>();
        public static Dictionary<string, double> lokal_ers_hpr = new Dictionary<string, double>();
        public static Dictionary<int, double> pengindex = new Dictionary<int, double>();
        public static Dictionary<int, double> takbelopp = new Dictionary<int, double>();
        public static int reference_year = 2023;
        public int endyear = -1;

        public static List<string> utbomrlist = new List<string>();
        public static Dictionary<string, string> utbomrdict = new Dictionary<string, string>();

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
            //string[] ss = Directory.GetDirectories(homefolder);
            //Directory.GetFiles(folder);

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
            instshortdict.Add("Institutionen för hälsa och välfärd", "IHOV");
            instshortdict.Add("Institutionen för information och teknik", "IIT");
            instshortdict.Add("Institutionen för kultur och samhälle", "IKS");
            instshortdict.Add("Institutionen för språk, litteratur och lärande", "ISLL");
            instshortdict.Add("Institutionen för lärarutbildning", "ILU");
            instshortdict.Add(utaninst, "?");
            instshortdict.Add(hda, hda);
            foreach (string inst in instshortdict.Keys)
                shortinstdict.Add(instshortdict[inst], inst);

            fill_subjinstdict();

            read_studentpeng(reference_year);
        }

        public void read_studentpeng(int refyear)
        {
            int n = 0;
            int year = DateTime.Now.Year -5;

            //Från budgetpropp 2024:
            //pengindex.Add(2025, 1.073);

            //Från budgetpropp 2025:
            pengindex.Add(2026, 1.0906);
            pengindex.Add(2027, 1.1084);

            string fn = folder + "\\ers_belopp_lokala YYYY.txt";

            Dictionary<int, double> sumpeng = new Dictionary<int, double>();

            do
            {
                year++;
                string fnyear = fn.Replace("YYYY", year.ToString());
                if (!File.Exists(fnyear))
                    continue;
                using (StreamReader sr = new StreamReader(fnyear))
                {
                    sr.ReadLine();
                    sr.ReadLine();
                    sr.ReadLine();
                    sumpeng.Add(year, 0);
                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        string[] words = line.Split('\t');
                        string area = words[0].Substring(words[0].IndexOf('(') + 1, 2);
                        double hstpeng = util.tryconvertdouble(words[1].Replace(" ", ""));
                        double hprpeng = util.tryconvertdouble(words[2].Replace(" ", ""));
                        sumpeng[year] += hstpeng + hprpeng;
                        if (year == refyear)
                        {
                            memo(area + "\t" + hstpeng + "\t" + hprpeng);
                            lokal_ers_hst.Add(area, hstpeng);
                            lokal_ers_hpr.Add(area, hprpeng);
                        }
                        n++;
                    }
                }
                //memo(n + " utbildningsområden i read_studentpeng");
                memo("Lokala ersättningsbelopp läst från " + fnyear);
            }
            while (year < DateTime.Now.Year + 2);

            foreach (int yr in sumpeng.Keys)
            {
                pengindex.Add(yr, sumpeng[yr] / sumpeng[refyear]);
            }

            for (int i = 2020; i < 2031; i++)
                memo(i.ToString()+":"+get_pengindex(i).ToString());

            utbomrdict.Add("FK Humanistiska området", "HU");
            utbomrdict.Add("FK Idrottsliga området", "ID");
            utbomrdict.Add("FK Juridiska området", "JU");
            utbomrdict.Add("FK Mediaområdet", "MM");
            utbomrdict.Add("FK Medicinska området", "ME");
            utbomrdict.Add("FK Naturvetenskapliga området", "NA");
            utbomrdict.Add("FK Samhällsvetenskapliga området", "SA");
            utbomrdict.Add("FK Tekniska området", "TE");
            utbomrdict.Add("FK Tekniska området (från mediaområdet)", "TE");
            utbomrdict.Add("FK Undervisningsområdet", "LU");
            utbomrdict.Add("FK Vårdområdet", "VÅ");
            utbomrdict.Add("FK Övriga områden", "ÖV");

            //ur budgetpropp 2025:
            takbelopp.Add(2023, 488276);
            takbelopp.Add(2024, 505826);
            takbelopp.Add(2025, 513824);
            takbelopp.Add(2026, 514089);
            takbelopp.Add(2027, 520417);
            takbelopp.Add(2028, 520417);
            takbelopp.Add(2029, 520417);
            takbelopp.Add(2030, 520417);
            takbelopp.Add(2031, 520417);
            takbelopp.Add(2032, 520417);


        }

        public static double get_pengindex(int year)
        {
            if (year <= 0)
                return pengindex[reference_year];

            if (year < 2000)
                return get_pengindex(2000 + year);

            if (pengindex.ContainsKey(year))
                return pengindex[year];

            if (year > pengindex.Keys.Max())
                return pengindex[pengindex.Keys.Max()];

            if (year < pengindex.Keys.Min())
                return pengindex[pengindex.Keys.Min()];

            return pengindex[reference_year];
        }

        public void fill_subjinstdict()
        {
            subjinstdict.Add("AB", "IKS");
            subjinstdict.Add("AR", "ISLL");
            subjinstdict.Add("AS", "IKS");
            subjinstdict.Add("AU", "IKS");
            subjinstdict.Add("BE", "?");
            subjinstdict.Add("BI", "ILU");
            subjinstdict.Add("BP", "ILU");
            subjinstdict.Add("BQ", "IKS");
            subjinstdict.Add("BY", "IIT");
            subjinstdict.Add("DT", "IIT");
            subjinstdict.Add("EG", "IIT");
            subjinstdict.Add("EN", "ISLL");
            subjinstdict.Add("ET", "IIT");
            subjinstdict.Add("EU", "IKS");
            subjinstdict.Add("FI", "IKS");
            subjinstdict.Add("FÖ", "IKS");
            subjinstdict.Add("FR", "ISLL");
            subjinstdict.Add("FY", "IIT");
            subjinstdict.Add("GG", "ILU");
            subjinstdict.Add("GT", "IIT");
            subjinstdict.Add("HI", "IKS");
            subjinstdict.Add("IE", "IIT");
            subjinstdict.Add("IH", "IHOV");
            subjinstdict.Add("IK", "IIT");
            subjinstdict.Add("IT", "ISLL");
            subjinstdict.Add("JP", "ISLL");
            subjinstdict.Add("KE", "IIT");
            subjinstdict.Add("KG", "IKS");
            subjinstdict.Add("KI", "ISLL");
            subjinstdict.Add("KT", "IIT");
            subjinstdict.Add("LI", "ISLL");
            subjinstdict.Add("LP", "IKS");
            subjinstdict.Add("MA", "IIT");
            subjinstdict.Add("MC", "IHOV");
            subjinstdict.Add("MD", "ILU");
            subjinstdict.Add("MI", "IIT");
            subjinstdict.Add("MÖ", "IIT");
            subjinstdict.Add("MP", "IIT");
            subjinstdict.Add("MT", "IIT");
            subjinstdict.Add("NA", "IKS");
            subjinstdict.Add("NV", "ILU");
            subjinstdict.Add("PA", "IKS");
            subjinstdict.Add("PE", "ILU");
            subjinstdict.Add("PG", "ILU");
            subjinstdict.Add("PR", "ISLL");
            subjinstdict.Add("PS", "IHOV");
            subjinstdict.Add("RK", "IKS");
            subjinstdict.Add("RV", "IKS");
            subjinstdict.Add("RY", "ISLL");
            subjinstdict.Add("SA", "IHOV");
            subjinstdict.Add("SK", "IKS");
            subjinstdict.Add("SO", "IKS");
            subjinstdict.Add("SP", "ISLL");
            subjinstdict.Add("SQ", "IIT");
            subjinstdict.Add("SR", "IHOV");
            subjinstdict.Add("SS", "ISLL");
            subjinstdict.Add("ST", "IIT");
            subjinstdict.Add("SV", "ISLL");
            subjinstdict.Add("SW", "IIT");
            subjinstdict.Add("TR", "IKS");
            subjinstdict.Add("TY", "ISLL");
            subjinstdict.Add("VÅ", "IHOV");
            subjinstdict.Add("VV", "IHOV");
            subjinstdict.Add("", hda);
            subjinstdict.Add("XX", hda);

            subjnamedict.Add("VÅ", "Omvårdnad");
            subjnamedict.Add("VV", "Vårdvetenskap");
            subjnamedict.Add("MC", "Medicinsk vetenskap");
            subjnamedict.Add("IH", "Idrott och hälsa");
            subjnamedict.Add("SR", "Sexuell, reproduktiv och perinatal hälsa");
            subjnamedict.Add("SA", "Socialt arbete");
            subjnamedict.Add("MI", "Mikrodata");
            subjnamedict.Add("IK", "Informatik");
            subjnamedict.Add("ST", "Statistik");
            subjnamedict.Add("EG", "Energiteknik");
            subjnamedict.Add("MT", "Maskinteknik");
            subjnamedict.Add("BY", "Byggteknik");
            subjnamedict.Add("DT", "Datateknik");
            subjnamedict.Add("ET", "Elektroteknik");
            subjnamedict.Add("EU", "Entreprenörskap");
            subjnamedict.Add("IE", "Industriell ekonomi");
            subjnamedict.Add("SK", "Statsvetenskap");
            subjnamedict.Add("AB", "Arbetsvetenskap");
            subjnamedict.Add("SO", "Sociologi");
            subjnamedict.Add("TR", "Turismvetenskap");
            subjnamedict.Add("FÖ", "Företagsekonomi");
            subjnamedict.Add("RK", "Religionsvetenskap");
            subjnamedict.Add("HI", "Historia");
            subjnamedict.Add("RV", "Rättsvetenskap");
            subjnamedict.Add("NA", "Nationalekonomi");
            subjnamedict.Add("LP", "Ljudproduktion");
            subjnamedict.Add("FI", "Filosofi");
            subjnamedict.Add("BP", "Bild");
            subjnamedict.Add("PE", "Pedagogik");
            subjnamedict.Add("PG", "Pedagogiskt arbete");
            subjnamedict.Add("NV", "Naturvetenskap");
            subjnamedict.Add("MD", "Matematikdidaktik");
            subjnamedict.Add("AR", "Arabiska");
            subjnamedict.Add("JP", "Japanska");
            subjnamedict.Add("SS", "Svenska som andraspråk");
            subjnamedict.Add("SP", "Spanska");
            subjnamedict.Add("IT", "Italienska");
            subjnamedict.Add("FR", "Franska");
            subjnamedict.Add("LI", "Litteraturvetenskap");
            subjnamedict.Add("SV", "Svenska");
            subjnamedict.Add("TY", "Tyska");
            subjnamedict.Add("EN", "Engelska");
            subjnamedict.Add("PR", "Portugisiska");
            subjnamedict.Add("RY", "Ryska");
            subjnamedict.Add("KI", "Kinesiska");
            subjnamedict.Add("FY", "Fysik");
            subjnamedict.Add("MA", "Matematik");
            subjnamedict.Add("KE", "Kemi");
            subjnamedict.Add("BQ", "Bildproduktion");
            subjnamedict.Add("PA", "Personal och arbetsliv");
            subjnamedict.Add("PS", "Psykologi");
            subjnamedict.Add("AS", "Afrikanska studier");
            subjnamedict.Add("AU", "Audiovisuella studier");
            subjnamedict.Add("BI", "Biologi");
            subjnamedict.Add("MÖ", "Miljö");
            subjnamedict.Add("GT", "Grafisk teknologi");
            subjnamedict.Add("KG", "Kulturgeografi");
            subjnamedict.Add("KT", "??");
            subjnamedict.Add("MP", "Materialteknik");
            subjnamedict.Add("SQ", "Samhällsbyggnad");

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
                    double?[] applvec = new double?[4] { null,null,null,null };
                    applvec[0] = appl;
                    applvec[3] = actualstud[0];
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

                    programbatchclass b = new programbatchclass(actualstud, origprogramdict[prog].id, words[batchcol], exam, applvec, res);
                    origprogramdict[prog].batchlist.Add(b);
                    nline++;
                }
            }
            memo("nline = " + nline);
        }

        private void read_retention_v3(string fn)
        {
            memo("Reading " + fn);
            int nline = 0;
            using (StreamReader sr = new StreamReader(fn))
            {
                string head = sr.ReadLine();
                sr.ReadLine();
                string[] hwords = head.Split('\t');
                int offset = 7;

                int progcol = 1;
                int batchcol = 0;
                int examcol = 2;
                int applcol = 3;
                int subjcol = -1;
                int sectorcol = -1;
                int reservecol = 7;
                int acceptcol = 4;
                int u1col = 5;
                int u2col = 6;

                for (int i=0;i<hwords.Length;i++)
                {
                    switch (hwords[i])
                    {
                        case "Kull": batchcol = i;
                            break;
                        case "Program":
                            progcol = i;
                            break;
                        case "Antal examinerade":
                            examcol = i;
                            break;
                        case " Behöriga 1a-handssökande":
                            applcol = i;
                            break;
                        case "Antagna nuläge":
                            acceptcol = i;
                            break;
                        case "Antagna urval 1":
                            u1col = i;
                            break;
                        case "Antagna urval 2":
                            u2col = i;
                            break;
                        case "Reserver":
                            reservecol = i;
                            break;
                        case "T1":
                            offset = i-1;
                            break;
                    }
                }

                //if (fn.Contains("_classified"))
                //{
                //    subjcol = 0;
                //    sectorcol = 1;
                //    progcol = 3;
                //    batchcol = 2;
                //    examcol = 4;
                //    applcol = 5;
                //    reservecol = 6;
                //    offset = 7;
                //}

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
                    actualstud[0] = util.tryconvertnull(words[acceptcol]);
                    int appl = util.tryconvert(words[applcol]);
                    if (appl < 0)
                        appl = 0; 
                    bool anynotnull = (actualstud[0] != null || appl > 0);
                    for (int i = 1; i < programbatchclass.maxsem; i++) 
                    {
                        if (offset + i >= words.Length || !hwords[offset+i].StartsWith("T"))
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

                    double?[] applvec = new double?[4] { null, null, null, null };
                    applvec[0] = appl;
                    applvec[1] = util.tryconvertnull(words[u1col]);
                    applvec[2] = util.tryconvertnull(words[u2col]);
                    applvec[3] = actualstud[0];
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

                    programbatchclass b = new programbatchclass(actualstud, origprogramdict[prog].id, words[batchcol], exam, applvec, res);
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
                    if (!programkoppling.ContainsKey(words[1]))
                    {
                        memo(words[1]);
                        continue;
                    }
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
            //string fn1 = folder + @"\programretention 230208.txt";
            //string fn1 = folder + @"\Programretention 230517-UHRcorrected.txt";
            //string fn1 = folder + @"\Programretention 230810.txt";
            //string fn1 = folder + @"\Programretention 231215.txt";
            string fn1 = folder + @"\Programretention 240229.txt";
            //string fn1 = folder + @"\Programretention 240830.txt";

            DateTime newest = DateTime.MinValue;
            foreach (string ffnn in Directory.GetFiles(folder))
            {
                //memo(ffnn);
                if (!ffnn.ToLower().Contains("\\programretention"))
                    continue;
                if (!ffnn.EndsWith(".txt"))
                    continue;
                if (File.GetCreationTime(ffnn) > newest)
                {
                    newest = File.GetCreationTime(ffnn);
                    fn1 = ffnn;
                }
            }
            //if (fn1.Contains("_v2"))
            read_retention_v3(fn1);
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

            if (year != reference_year)
            {
                memo("=====================================================");
                memo("Fel reference_year! Fixa i Form1 header");
                memo("så att det blir samma som 'prod per inst och prog'");
                memo("=====================================================");
                Console.ReadLine();
            }

            programclass.prodyear = year;
            read_prod(fnyear);


            foreach (string p in origprogramdict.Keys)
            {
                origprogramdict[p].normalize_per_student();
                memo(p + "\t" + origprogramdict[p].prod_per_student.frachst.ToString("N1") + " hst per stud");
                //if (p == "Produktionstekniker 120 hp")
                //    origprogramdict[p].fracproddict = origprogramdict["Energiteknikerprogrammet"].fracproddict;
                if (p == "Ekonomprogrammet - Business Management")
                    origprogramdict[p].fracproddict = origprogramdict["Ekonomprogrammet"].fracproddict;
                if (p == "Filmproduktion 180 hp")
                    origprogramdict[p].fracproddict = origprogramdict["Film- och TV produktion"].fracproddict;
                if (p == "Manus och producentskap 180 hp")
                    origprogramdict[p].fracproddict = origprogramdict["Film- och TV produktion"].fracproddict;
                if (p == "Medieproduktion för TV och sociala medier 180 hp")
                    origprogramdict[p].fracproddict = origprogramdict["Film- och TV produktion"].fracproddict;
                if (p == "Musik- och ljudproduktion 180 hp")
                    origprogramdict[p].fracproddict = origprogramdict["Film- och TV produktion"].fracproddict;
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
                    double apptrans = pc.appltransition[0][1].transitionprob;
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
                        int iapp = findlimitindex((double)pb.applicants[0], applimit);
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

        private List<string> parseapplcodes(string input)
        {
            List<string> ls = new List<string>();

            string rex = @"[HV][\w\d]{4}";

            foreach (Match m in Regex.Matches(input, rex))
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

        public static programclass findprogrambyname(string name)
        {
            var q = from c in origprogramdict.Values where c.name.ToLower() == name.ToLower() select c;
            return q.FirstOrDefault();
        }

        public static programclass findprogram(string code)
        {
            string c2 = code;
            if (programkoppling.ContainsKey(code))
                c2 = programkoppling[code];
            return findcourse(c2, origprogramdict,new Dictionary<string, programclass>());
        }

        public static programclass findcourse(string code)
        {
            return findcourse(code, fkdict,fkcodedict);
        }

        public static programclass findcourse(string code, Dictionary<string,programclass> cdict, Dictionary<string, programclass> codedict) //either code or name as input
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
                            fk.activecourse = true;
                            fk.subjectcode = getsubjectcode(coursecode);
                            if (subjinstdict.ContainsKey(fk.subjectcode))
                                fk.homeinst = shortinstdict[subjinstdict[fk.subjectcode]];
                            if (!String.IsNullOrEmpty(words[10]))
                            {
                                string pack = words[10].Trim(',');
                                if (!fk.partofpackage.Contains(pack))
                                    fk.partofpackage.Add(pack);
                            }
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

            //foreach (string s in fkdict.Keys)
            //{
            //    fkdict[s].calculate_transitions();
            //}

            //var qpart = from c in fkdict.Values
            //            where !String.IsNullOrEmpty(c.partofpackage)
            //            select c;
            //foreach (programclass part in qpart)
            //{
            //    programclass paket = findprogram(part.partofpackage);
            //    if (paket == null)
            //    {
            //        memo(part.partofpackage + " not found");
            //    }
            //    else
            //    {
            //        paket.homeinst = part.homeinst;
            //        if (!paket.coursedict.ContainsKey(1))
            //            paket.coursedict.Add(1, new Dictionary<string, double>());
            //        paket.coursedict[1].Add(part.bestcode(), 1);
            //    }
            //}

        }



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
                    string header = sr.ReadLine(); //header line

                    string[] hwords = header.Split('\t');

                    int subjectcol = -1;
                    int sectorcol = -1;
                    int semcol = -1;
                    int namecol = -1;
                    int codecol = -1;
                    int applcol = -1;
                    int acceptcol = -1;
                    int u1col = -1;
                    int u2col = -1;
                    int regcol = -1;
                    int examcol = -1;
                    int tillfallecol = -1;

                    if ( header.StartsWith("Subject"))
                    {
                        subjectcol = 0;
                        sectorcol = 1;
                        semcol = 2;
                        codecol = 3;
                        namecol = 4;
                        applcol = 5;
                        acceptcol = 6;
                        regcol = 7;
                        examcol = 8;
                    }
                    else if (header.Contains("urval"))
                    {
                        semcol = 0;
                        codecol = 1;
                        namecol = 2;
                        applcol = 3;
                        acceptcol = 4;
                        u1col = 5;
                        u2col = 6;
                        regcol = 7;
                        examcol = 8;
                        tillfallecol = 9;
                    }

                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        string[] words = line.Split('\t');
                        double hp = -1;
                        //string name = words[namecol];
                        string code = words[codecol];
                        var hpresult = util.extract_hp(words[namecol]);
                        string name = hpresult.Item1;
                        hp = hpresult.Item2;
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
                            if (subjectcol > 0)
                            {
                                fk.subject = words[subjectcol];
                                fk.sector = words[sectorcol];
                            }
                            fkdict.Add(name, fk);
                        }
                        //public programbatchclass(double?[] actualstud, int prog, string bstart, int exam, int appl, int res) //real data
                        double?[] actualstud = new double?[2];
                        actualstud[0] = util.tryconvert0(words[acceptcol]);
                        actualstud[1] = util.tryconvert0(words[regcol]);
                        double?[] applvec = new double?[4] { null, null, null, null };
                        applvec[0] = util.tryconvert0(words[applcol]);
                        if (u1col > 0)
                        {
                            applvec[1] = util.tryconvert0(words[u1col]);
                            applvec[2] = util.tryconvert0(words[u2col]);
                        }
                        applvec[3] = actualstud[0];
                        programbatchclass kt = new programbatchclass(actualstud, -1, util.semester3to2(words[semcol]), util.tryconvert(words[examcol]), applvec, 0);
                        fk.batchlist.Add(kt);
                        if (!fk.coursecodelist.Contains(words[codecol]))
                        {
                            fk.coursecodelist.Add(words[codecol]);
                            fk.subjectcode = getsubjectcode(words[codecol]);
                            fk.homeinst = (shortinstdict[subjinstdict[fk.subjectcode]]);
                        }
                        if (!fkcodedict.ContainsKey(words[codecol]))
                            fkcodedict.Add(words[codecol], fk);

                    }
                }
                memo("# courses = " + fkdict.Count);
            }

            //foreach (string p in fkdict.Keys)
            //    fkdict[p].calculate_transitions();

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

        private void read_pccourseincomedict(string fn)
        {
            memo("Reading " + fn);

            List<string> progs = new List<string>();

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
                    string prog = words[2];
                    string code = words[1].Substring(0,6);
                    if (string.IsNullOrEmpty(code))
                        continue;
                    string coursename = words[1].Replace(code,"").Trim();
                    double hp = -1;
                    if (coursename.Contains("hp"))
                    {
                        var hpresult = util.extract_hp(coursename);
                        coursename = hpresult.Item1;
                        hp = hpresult.Item2;
                    }
                    double hst = util.tryconvertdouble(words[3]);
                    double hpr = util.tryconvertdouble(words[4]);
                    double income = util.tryconvertdouble(words[7]);
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
                            //if (hp == course.hp || course.hp < 0 || hp < 0)
                            {
                                nfoundname++;
                                course.coursecodelist.Add(code);
                                course.hp = hp;
                                fkcodedict.Add(code, course);
                            }
                            //else
                            //    course = null;
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
                        course.hp = hp;
                        //course.subject = words[0];
                        //course.sector = words[1];
                        course.coursecodelist.Add(code);
                        course.subjectcode = getsubjectcode(code);
                        course.homeinst = (shortinstdict[subjinstdict[course.subjectcode]]);
                        var c2 = (from c in fkdict.Values
                                  where c.subjectcode == course.subjectcode
                                  select c).FirstOrDefault();
                        if (c2 != null)
                            course.studentpengarea = c2.studentpengarea;
                        fkdict.Add(coursename, course);
                        fkcodedict.Add(code, course);
                        //continue;
                    }
                    string bestcode = course.bestcode();
                    //public Dictionary<string, double> coursehstdict = new Dictionary<string, double>();
                    //public Dictionary<string, double> courseincomedict = new Dictionary<string, double>();
                    //public Dictionary<string, double> coursecostdict = new Dictionary<string, double>();
                    if (!pc.coursehstdict.ContainsKey(bestcode))
                    {
                        pc.coursehstdict.Add(bestcode, hst);
                        pc.courseincomedict.Add(bestcode, income);
                        pc.coursecostdict.Add(bestcode, 0);
                    }
                    else
                    {
                        pc.coursehstdict[bestcode] += hst;
                        pc.courseincomedict[bestcode] += income;
                    }

                }
                memo("# lines " + nline);
                memo("# progs = " + progs.Count);
                memo("#courses found by code: " + nfoundcode);
                memo("#courses found by name: " + nfoundname);
                memo("#courses not found: " + nnfound);
                memo("#programs not found: " + nprognotfound);

            }

        }


        private void read_pccoursedict(string fn,bool activate)
        {
            List<string> progs = new List<string>();

            int nfoundcode = 0;
            int nfoundname = 0;
            int nnfound = 0;
            int nprognotfound = 0;
            if (!File.Exists(fn))
            {
                memo("### File not found: " + fn);
                return;
            }
            else
                memo("Reading " + fn);
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
                    double hp = -1;
                    if (coursename.Contains("hp"))
                    {
                        var hpresult = util.extract_hp(coursename);
                        coursename = hpresult.Item1;
                        hp = hpresult.Item2;
                    }
                    int sem = util.tryconvert(words[3]);
                    double stud = util.tryconvertdouble(words[4]);
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
                            //if (hp == course.hp || course.hp < 0 || hp < 0)
                            {
                                nfoundname++;
                                course.coursecodelist.Add(code);
                                course.hp = hp;
                                fkcodedict.Add(code, course);
                            }
                            //else
                            //    course = null;
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
                        course.hp = hp;
                        //course.subject = words[0];
                        //course.sector = words[1];
                        course.coursecodelist.Add(code);
                        course.subjectcode = getsubjectcode(code);
                        course.homeinst = (shortinstdict[subjinstdict[course.subjectcode]]);
                        var c2 = (from c in fkdict.Values
                                 where c.subjectcode == course.subjectcode
                                 select c).FirstOrDefault();
                        if (c2 != null)
                            course.studentpengarea = c2.studentpengarea;
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
                    if (activate)
                        fkcodedict[bestcode].activecourse = true;
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

        private void read_program_programkurser_intag()
        {
            openFileDialog1.Title = "Select Program_programkurser_intag file";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fn = openFileDialog1.FileName;
                memo("Reading program/course data from " + fn);

                read_pccoursedict(fn,false);
            }
        }

        private void read_lokal_ers_programkurser()
        {
            openFileDialog1.Title = "Select Lokal_ersättning_programkurser file";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fn = openFileDialog1.FileName;
                memo("Reading program/course-income data from " + fn);

                read_pccourseincomedict(fn);
            }
        }

        private void read_special_pccoursedict()
        {
            string fn = folder + @"\special_pccoursedict_"+DateTime.Now.Year+".txt";
            memo("Reading w/o normalizing " + fn);
            read_pccoursedict(fn,true);
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

            read_special_pccoursedict();

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

                if (!utbomrlist.Contains(s4[0]))
                    utbomrlist.Add(s4[0]);
            }
            return dict;
        }

        public static double hstkr(double hst, Dictionary<string,double> hstpeng, int year)
        {
            double kr = 0;
            foreach (string area in hstpeng.Keys)
            {
                kr += lokal_ers_hst[area] * hstpeng[area];
            }
            //if (pengindex.ContainsKey(year))
            if (year > 0)
                kr *= get_pengindex(year);
            return kr * hst;
        }

        public static double hprkr(double hst, Dictionary<string, double> hstpeng,int year)
        {
            double kr = 0;
            foreach (string area in hstpeng.Keys)
            {
                kr += lokal_ers_hpr[area] * hstpeng[area];
            }
            if (year > 0)
                kr *= get_pengindex(year);
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
                    double krhst = hstkr(hst, course.studentpengarea,-1);
                    double krhpr = hprkr(hpr, course.studentpengarea,-1);
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
            int i = 2019;
            string fn = folder + @"\hst_hpr_utfall_budget_reg " + i + ".txt";
            while (File.Exists(fn))
            {
                var dict = read_hst_hpr(fn);
                do_hst_hpr_dict(dict);
                i++;
                fn = folder + @"\hst_hpr_utfall_budget_reg " + i + ".txt";
            }
            string fnmiss = folder + @"\missing-courses.txt";
            var dmiss = read_hst_hpr(fnmiss);
            do_hst_hpr_dict(dmiss);
        }

        private void coursedatabutton_Click(object sender, EventArgs e)
        {
            read_aktiva_utb_file();
            read_fkfile();
            read_program_programkurser_intag();
            read_antagningsstatistik_linnea(folder + "\\" + "antagning kurspaket V22 per_utb_v2.txt");
            read_antagningsstatistik_linnea(folder + "\\" + "antagning kurspaket H22 per_utb_v2.txt");
            read_antagningsstatistik_linnea(folder + "\\" + "antagning kurspaket V23 per_utb_v2.txt");
            read_antagningsstatistik_linnea(folder + "\\" + "antagning kurspaket H23 per_utb_v2.txt");
            read_antagningsstatistik_linnea(folder + "\\" + "antagning kurspaket V24 per_utb_v2.txt");
            read_antagningsstatistik_linnea(folder + "\\" + "antagning kurspaket H24 per_utb_v2.txt");
            fk_transitions_parts();
            do_hst_hpr();
            normalize_pccoursedict();
            read_kurspaketregistrering("VT22"); //reads from VT22 onwards as long as there is data

            read_exchangestudents(2023);

            fill_triangledicts();

            //fixa program som saknar data:
            foreach (string p in origprogramdict.Keys)
            {
                origprogramdict[p].extendretention(0.7);
                origprogramdict[p].set_homeinst();

                if (origprogramdict[p].utype == "Program" && origprogramdict[p].homeinst != utaninst && origprogramdict[p].fracproddict.Count == 0)
                {
                    if (origprogramdict[p].semesters > 1)
                    {
                        var qprog = (from c in origprogramdict.Values
                                    where c.homeinst == origprogramdict[p].homeinst
                                    where c.utype == "Program"
                                    where c.semesters == origprogramdict[p].semesters
                                    where c.fracproddict.Count > 0
                                    select c).ToList();
                        if (qprog.Count == 0)
                        {
                            qprog = (from c in origprogramdict.Values
                                     where c.homeinst == origprogramdict[p].homeinst
                                     where c.utype == "Program"
                                     //where c.semesters == origprogramdict[p].semesters
                                     where c.fracproddict.Count > 0
                                     select c).ToList();

                        }
                        var pclone = programclass.clone(qprog);
                        origprogramdict[p].datafromtemplate(pclone);
                    }
                }
            }

            Dictionary<string, Dictionary<string, double>> subjectpengdict = new Dictionary<string, Dictionary<string, double>>();

            //fixa saknade studentpengar i kurser:
            foreach (programclass pc in fkdict.Values)
            {
                if (pc.studentpengarea.Count > 0)
                {
                    if (!subjectpengdict.ContainsKey(pc.subjectcode))
                        subjectpengdict.Add(pc.subjectcode, pc.studentpengarea);
                }
            }
            foreach (programclass pc in fkdict.Values)
            {
                if (pc.studentpengarea.Count == 0)
                {
                    if (subjectpengdict.ContainsKey(pc.subjectcode))
                        pc.studentpengarea = new Dictionary<string, double>(subjectpengdict[pc.subjectcode]);
                }
            }


            Excelplanningbutton.Enabled = true;
            memo("==== DONE ====");
        }

        private void read_exchangestudents(int startyear)
        {
            string fnbase = folder + "\\kursregistrering utbytesavtal ";
            string fn = fnbase + startyear + ".txt";
            int year = startyear;
            while (File.Exists(fn))
            {
                int nstud = 0;
                int ncourse = 0;
                memo("Reading " + fn);
                using (StreamReader sr = new StreamReader(fn))
                {
                    sr.ReadLine();
                    sr.ReadLine();
                    sr.ReadLine();

                    string hline = sr.ReadLine();
                    string[] hwords = hline.Split('\t');

                    string currentcode = "";
                    programclass currentcourse = null;

                    while (!sr.EndOfStream)
                    {
                        string line1 = sr.ReadLine();
                        var codelist = parsecoursecodes(line1);
                        if (codelist.Count > 0)
                        {
                            if (codelist.First() != currentcode)
                            {
                                currentcode = codelist.First();
                                string name = line1.Replace(currentcode, "").Trim().Trim('"');
                                currentcourse = findcourse(currentcode);
                                if (currentcourse == null)
                                    currentcourse = findcourse(name);
                                if (currentcourse == null)
                                {
                                    currentcourse = new programclass(name);
                                    currentcourse.name = name;
                                    fkdict.Add(name, currentcourse);
                                }
                                if (!fkcodedict.ContainsKey(currentcode))
                                    fkcodedict.Add(currentcode, currentcourse);
                                if (!currentcourse.coursecodelist.Contains(currentcode))
                                {
                                    currentcourse.coursecodelist.Add(currentcode);
                                    currentcourse.subjectcode = getsubjectcode(currentcode);
                                    currentcourse.homeinst = (shortinstdict[subjinstdict[currentcourse.subjectcode]]);
                                }
                            }
                        }
                        else if (line1.Contains("tbytesavtal"))
                        {
                            string line2 = sr.ReadLine();
                            string[] w1 = line1.Split('\t');
                            string[] w2 = line2.Split('\t');
                            if (parseapplcodes(w1[0]).Count > 0)
                            {
                                string applcode = parseapplcodes(w1[0]).First();
                                currentcourse.applcodelist.Add(applcode);
                                string batch = applcode[0] + "T" + (year % 100);
                                var bc = currentcourse.getbatch(batch);
                                if (bc != null)
                                {
                                    int exstud = util.tryconvert(w1[2]);
                                    if (exstud <= 0)
                                        exstud = util.tryconvert(w2[2]);
                                    if (exstud > 0)
                                    {
                                        bc.exchangestud = exstud;
                                        nstud += exstud;
                                        ncourse++;
                                    }
                                }
                            }

                        }
                    }
                    memo("Exchange students " + nstud + " in " + ncourse + " courses");
                }
                year++;
                fn = fnbase + year + ".txt";
            }
            //if (!fkcodedict.ContainsKey(applcode))
            //    fkcodedict.Add(applcode, fk);
            //if (!fk.applcodelist.Contains(applcode))
            //{
            //    fk.applcodelist.Add(applcode);
            //}
        }

        string parenrex = @"\((.*)\)";
        string daterex = @" (\d+ \w+) ";

        private void read_antagningsstatistik_linnea(string fn)
        {
            int u1col = 3;
            int applindex = 2;
            //hbookclass typehist = new hbookclass("Utbildningstyp");
            //openFileDialog1.InitialDirectory = folder;
            //openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            //openFileDialog1.Title = "Select antagningsstatistik per utb file";
            //Console.WriteLine("opendialog1.Show:");
            //if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //string fn = openFileDialog1.FileName;
                memo("Reading antagningsstatistik from " + fn);
                using (StreamReader sr = new StreamReader(fn))
                {
                    string dateline = sr.ReadLine();
                    string yearline = sr.ReadLine();
                    string yearstring = yearline.Split()[1].Trim();
                    foreach (Match m in Regex.Matches(dateline,daterex+yearstring))
                    {
                        string date = m.Groups[1].Value;
                        if (date.Contains("juli"))
                            applindex = 1;
                        else
                            applindex = 2;
                    }
                    string semline = sr.ReadLine();
                    string sem = util.semester3to2(semline.Split()[1].Trim('"'));
                    sr.ReadLine(); //throw away header line

                    List<string> alreadyu1 = new List<string>();
                    foreach (programclass pc in origprogramdict.Values)
                    {
                        if (pc.getbatch(sem) != null && pc.getbatch(sem).applicants[1] != null)
                            alreadyu1.Add(pc.name);
                    }
                    foreach (programclass pc in fkdict.Values)
                    {
                        if (pc.getbatch(sem) != null && pc.getbatch(sem).applicants[1] != null)
                            alreadyu1.Add(pc.name);
                    }


                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        if (line.ToUpper().Contains("INSTÄLLD"))
                            continue;
                        string[] words = line.Split('\t');
                        if (String.IsNullOrEmpty(words[0].Trim()))
                            continue;
                        string ss = words[0].Trim();
                        if (ss == "Utbildning")
                            continue;
                        double hp = util.tryconvertdouble(ss.Split().Last().Replace("hp",""));
                        string name = ss.Split('(').First().Trim();
                        if (name.StartsWith("Senare del"))
                            continue;
                        string applcode = "";
                        string distnorm = "";

                        var mm = Regex.Matches(ss, parenrex);
                        if (mm.Count > 0)
                            applcode = mm[0].Groups[1].Value;
                        if (mm.Count > 1)
                            distnorm = mm[1].Groups[1].Value;

                        string coursecode = words[1];
                        string utype = String.IsNullOrEmpty(coursecode) ? "Program" : "Kurs";
                        if (utype == "Program" && hp <= 50)
                            utype = "Kurspaket";
                        //typehist.Add(utype);
                        //double fullfee = util.tryconvertdouble(words[16].Replace(" ", ""));

                        int? accepted = util.tryconvert(words[3]);
                        if (accepted <= 0)
                            accepted = null;

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
                            else if (alreadyu1.Contains(fk.name))
                                continue;
                            if (fk.hp <= 0)
                                fk.hp = hp;
                            fk.semesters = 1;
                            fk.utype = utype;
                            //fk.fee = fullfee;
                            fk.fk = true;
                            fk.subjectcode = getsubjectcode(coursecode);
                            if (subjinstdict.ContainsKey(fk.subjectcode))
                                fk.homeinst = shortinstdict[subjinstdict[fk.subjectcode]];
                            //if (!String.IsNullOrEmpty(words[10]))
                            //    fk.partofpackage = words[10].Trim(',');
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
                            if (fk.getbatch(sem) == null)
                            {
                                double?[] actualstud = new double?[2] { null,null };
                                double?[] applvec = new double?[4] { null, null, null, null };
                                //applvec[0] = util.tryconvert(words[applcol]);
                                applvec[applindex] = accepted;
                                applvec[3] = actualstud[0];
                                programbatchclass kt = new programbatchclass(actualstud, -1, util.semester3to2(sem), 0, applvec, 0);
                                fk.batchlist.Add(kt);
                            }
                            else
                            {
                                var bc = fk.getbatch(sem);
                                if (bc.applicants[applindex] == null)
                                    bc.applicants[applindex] = accepted;
                                else
                                    bc.applicants[applindex] += accepted;
                            }
                        }
                        else
                        {
                            programclass pc = findprogram(name);
                            if (pc == null && !String.IsNullOrEmpty(applcode))
                                pc = findprogram(applcode);
                            if (pc == null)
                            {
                                if (name.Contains("mneslärar"))
                                {
                                    if (name.Contains("ymnasi"))
                                        name = "Ämneslärare Gymnasieskolan";
                                    else if (name.Contains("grundskolan"))
                                        name = "Ämneslärare - Grundskolans årskurs 7-9";
                                    pc = findprogram(name);
                                }
                            }
                            if (pc == null) //create new entry
                            {
                                memo("Program not found " + name);
                                pc = new programclass(name);
                                pc.name = name;
                                origprogramdict.Add(name, pc);
                            }
                            else if (alreadyu1.Contains(pc.name))
                                continue;
                            if (pc.hp <= 0)
                                pc.hp = hp;
                            pc.semesters = (int)Math.Ceiling(hp / 30);
                            pc.utype = utype;
                            //pc.fee = fullfee;
                            pc.fk = false;
                            if (!String.IsNullOrEmpty(applcode) && !pc.applcodelist.Contains(applcode))
                            {
                                pc.applcodelist.Add(applcode);
                            }
                            if (pc.getbatch(sem) == null)
                            {
                                double?[] actualstud = new double?[programbatchclass.maxsem];
                                for (int i = 0; i < programbatchclass.maxsem; i++)
                                    actualstud[i] = null;
                                double?[] applvec = new double?[4] { null, null, null, null };
                                //applvec[0] = util.tryconvert(words[applcol]);
                                applvec[applindex] = accepted;
                                applvec[3] = actualstud[0];
                                programbatchclass kt = new programbatchclass(actualstud, -1, util.semester3to2(sem), 0, applvec, 0);
                                pc.batchlist.Add(kt);
                            }
                            else
                            {
                                var bc = pc.getbatch(sem);
                                if (bc.applicants[applindex] == null)
                                    bc.applicants[applindex] = accepted;
                                else
                                    bc.applicants[applindex] += accepted;
                            }
                        }
                    }
                }
                memo("# courses = " + fkdict.Count);
            }

            //memo(typehist.GetSHist());


        }

        private void fk_transitions_parts()
        {
            foreach (string s in fkdict.Keys)
            {
                fkdict[s].calculate_transitions();
            }

            var qpart = from c in fkdict.Values
                        where c.partofpackage.Count > 0
                        select c;
            foreach (programclass part in qpart)
                foreach (string pack in part.partofpackage)
                {
                    programclass paket = findprogram(pack);
                    if (paket == null)
                    {
                        memo(pack + " not found");
                    }
                    else
                    {
                        paket.homeinst = part.homeinst;
                        if (!paket.coursedict.ContainsKey(1))
                            paket.coursedict.Add(1, new Dictionary<string, double>());
                        if (!paket.coursedict[1].ContainsKey(part.bestcode()))
                            paket.coursedict[1].Add(part.bestcode(), 1);
                    }
                }
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

        private void read_kurspaketregistrering(string startsem)
        {
            string fnbase = folder + "\\kurspaketregistrering ";
            string sem = startsem;
            string endstring = " Total";
            while (File.Exists(fnbase+sem+".txt"))
            {
                using (StreamReader sr = new StreamReader(fnbase + sem + ".txt"))
                {
                    memo("Reading " + fnbase + sem + ".txt");
                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        string[] words = line.Split('\t');
                        if (words[0].EndsWith(endstring) && words.Length > 2)
                        {
                            string name = words[0].Replace(endstring, "").Trim();
                            var tup = util.extract_hp(name);
                            name = tup.Item1;
                            programclass paket = findprogrambyname(name);
                            if (paket != null)
                            {
                                int reg = util.tryconvert0(words[2]);
                                //memo(name + "\t" + sem + "\t" + reg);
                                programbatchclass pb = paket.getbatch(sem);
                                if (pb == null)
                                {
                                    double?[] actualstud = new double?[] { null, reg };
                                    pb = new programbatchclass(actualstud, paket.id, sem);
                                    paket.batchlist.Add(pb);
                                }
                                else
                                {
                                    pb.setstud(reg, 1);
                                }
                            }
                            else
                            {
                                memo(name + " not found");
                            }
                        }
                    }
                }
                sem = util.incrementsemester(sem);
            }
        }

        private void retention_by_institution()
        {
            int startyear = 21;
            int endyear = 23;
            StringBuilder sbhead = new StringBuilder("Inst");
            for (int year = startyear; year <= endyear; year++)
                sbhead.Append("\t" + year);
            memo(sbhead.ToString());
            double[] t1 = new double[endyear - startyear + 1];
            double[] t2 = new double[endyear - startyear + 1];
            foreach (string inst in institutiondict.Keys)
            {
                StringBuilder sb = new StringBuilder(inst);
                var q = from c in origprogramdict.Values
                        where c.homeinst == inst
                        select c;
                for (int year = startyear; year <= endyear; year++)
                {
                    double t1stud = 0;
                    double t2stud = 0;
                    string htstart = "HT" + (year - 1);
                    string vtstart = "VT" + year;
                    foreach (programclass pc in q)
                    {
                        programbatchclass pbht = pc.getbatch(htstart);
                        if (pbht != null && pbht.getactualstud(2) != null && pbht.getactualstud(1) != null)
                        {
                            t1stud += (double)pbht.getactualstud(1);
                            t2stud += (double)pbht.getactualstud(2);
                        }
                        programbatchclass pbvt = pc.getbatch(vtstart);
                        if (pbvt != null && pbvt.getactualstud(2) != null && pbvt.getactualstud(2) != null)
                        {
                            t1stud += (double)pbvt.getactualstud(1);
                            t2stud += (double)pbvt.getactualstud(2);
                        }
                        //memo("\t"+pc.name + "\t"+year+"\t" + t1stud + "\t" + t2stud);
                    }
                    sb.Append("\t" + (100 * t2stud / t1stud).ToString("F1") + " %");
                    t1[year - startyear] += t1stud;
                    t2[year - startyear] += t2stud;

                }
                memo(sb.ToString());

            }
            StringBuilder sbhda = new StringBuilder("HDa");
            for (int year = startyear; year <= endyear; year++)
                sbhda.Append("\t" + (100 * t2[year - startyear] / t1[year - startyear]).ToString("F1") + " %");
            memo(sbhda.ToString());

        }

        private void appl_to_T3()
        {
            double sumappl = 0;
            double sumT3 = 0;

            memo("Program\tSökande\tT3-stud\tSökande/T3");

            foreach (programclass pc in origprogramdict.Values)
            {
                double sumprogappl = 0;
                double sumprogT3 = 0;
                foreach (programbatchclass pb in pc.batchlist)
                {
                    if (pb.applicants != null && pb.applicants[0] != null)
                    {
                        if (pb.applicants[0] > 0 && pb.reserves == 0)
                        {
                            if (pb.getactualstud(3) != null && pb.getactualstud(3) > 0)
                            {
                                sumprogappl += (double)pb.applicants[0];
                                sumprogT3 += (double)pb.getactualstud(3);
                            }
                        }
                    }
                    sumappl += sumprogappl;
                    sumT3 += sumprogT3;
                }
                if (sumprogT3 > 0)
                    memo(pc.name + "\t" + sumprogappl + "\t" + sumprogT3 + "\t" + sumprogappl / sumprogT3);
            }
            memo("");
            memo("Totalt\t" + sumappl + "\t" + sumT3 + "\t" + sumappl / sumT3);
        }

        private void Retentionbutton_Click(object sender, EventArgs e)
        {
            retention_by_institution();
            appl_to_T3();
        }

        private void applicantbutton_Click(object sender, EventArgs e)
        {
            string startsem = "HT21";
            string endsem = "VT24";
            string refsem = "HT23";
            memo("\t\t\tSökande 1:ahand\t\t\t\t\t\t\t\t\tStudenter per programtermin HT23");
            StringBuilder sbhead = new StringBuilder("Programt\tInst\tÄmne");
            Dictionary<string, int> appldict = new Dictionary<string, int>();
            
            string sem = startsem;
            while (!util.comparesemesters(endsem,sem))
            {
                appldict.Add(sem, 0);
                sbhead.Append("\t" + sem);
                sem = util.incrementsemester(sem);
            }
            sbhead.Append("\tMedelsök\t% reg av antagna\tMedelreserver");
            for (int i=1;i<12;i++)
            {
                sbhead.Append("\tT" + i);
            }
            memo(sbhead.ToString());
            foreach (programclass pc in origprogramdict.Values.OrderBy(c=>c.name))
            {
                if (pc.batchlist.Count == 0)
                    continue;
                StringBuilder sb = new StringBuilder(pc.name+"\t"+instshortdict[pc.homeinst]);

                Dictionary<string, double> progsubjdict = new Dictionary<string, double>();
                foreach (int t in pc.coursedict.Keys)
                {
                    foreach (string code in pc.coursedict[t].Keys)
                    {
                        string subj = getsubjectcode(code);
                        if (!progsubjdict.ContainsKey(subj))
                            progsubjdict.Add(subj, 0);
                        progsubjdict[subj] += pc.coursedict[t][code];
                    }
                }
                string progsubj = "(none)";
                double max = -1;
                foreach (string subj in progsubjdict.Keys)
                {
                    if (progsubjdict[subj] > max)
                    {
                        max = progsubjdict[subj];
                        progsubj = subj;
                    }
                }
                sb.Append("\t" + progsubj);

                Dictionary<int, int> semstuddict = new Dictionary<int, int>();

                double sumappl = 0;
                int nbappl = 0;
                int nbreg = 0;
                double sumacc = 0;
                double sumreg = 0;
                double sumres = 0;
                bool foundbatch = false;
                foreach (programbatchclass pb in pc.batchlist)
                {
                    if (!util.comparesemesters(pb.batchstart,startsem) && !util.comparesemesters(endsem, pb.batchstart))
                    {
                        nbappl++;
                        if (pb.applicants[0] != null)
                        {
                            sumappl += (int)pb.applicants[0];
                            appldict[pb.batchstart] = (int)pb.applicants[0];
                        }
                        if (pb.actualsemstud != null 
                            && pb.actualsemstud[1] != null 
                            && pb.applicants[3] != null)
                        {
                            nbreg++;
                            sumacc += (double)pb.applicants[3];
                            sumreg += (double)pb.actualsemstud[1];
                            sumres += pb.reserves;
                        }

                    }
                    int tref = util.semestercount(pb.batchstart, refsem);
                    int refstud = (int)pb.getstud(tref);
                    if (refstud > 0)
                        semstuddict.Add(tref, refstud);
                }
                foreach (string sm in appldict.Keys.ToList())
                {
                    if (appldict[sm] > 0)
                    {
                        sb.Append("\t" + appldict[sm]);
                        foundbatch = true;
                    }
                    else
                        sb.Append("\t");
                    appldict[sm] = 0;
                }
                if (nbappl > 0)
                    sb.Append("\t" + (sumappl / nbappl).ToString("N1"));
                else
                    sb.Append("\t");
                if (nbreg > 0)
                {
                    sb.Append("\t" + (100 * sumreg / sumacc).ToString("N1")+"%");
                    sb.Append("\t" + (sumres / nbreg).ToString("N1"));
                }
                else
                    sb.Append("\t");
                for (int i=1;i<12;i++)
                {
                    if (semstuddict.ContainsKey(i))
                        sb.Append("\t" + semstuddict[i]);
                    else
                        sb.Append("\t");
                }
                if (foundbatch)
                    memo(sb.ToString());
            }
            memo("DONE");
        }

        private void programprofitbutton_Click(object sender, EventArgs e)
        {
            var q = from c in origprogramdict.Values where c.courseincomedict.Count() > 0 select c;
            if (q.Count() == 0)
            {
                read_lokal_ers_programkurser();
                read_kurskostnad();
            }

            memo("Namn\tHST\tIntäkt\tLärarkostnad\tTB\tInkl OH\t'Vinst'\tKostnad T1\tT2\tT3\tT4\tT5\tT6");
            foreach (programclass pc in origprogramdict.Values)
            {
                if (pc.courseincomedict.Count() == 0)
                    continue;
                StringBuilder sb = new StringBuilder(pc.name);
                double sumhst = 0;
                double sumincome = 0;
                double sumcost = 0;
                double[] semcost = new double[7] {0,0,0,0,0,0,0};
                foreach (string code in pc.courseincomedict.Keys)
                {
                    sumhst += pc.coursehstdict[code];
                    programclass course = findcourse(code);
                    if (course == null || course.coursehstdict.Count() == 0)
                    {
                        memo("\t"+code+" missing data");
                        continue;
                    }
                    string bestcode = course.bestcode();
                    double progfraction = course.courseincomedict[bestcode] > 0 ? pc.courseincomedict[bestcode] / course.courseincomedict[bestcode] : 0;
                    double courseprogcost = progfraction * course.coursecostdict[bestcode];
                    sumincome += progfraction * course.courseincomedict[bestcode];
                    sumcost += courseprogcost;
                    //memo("\t" + bestcode + "\t" + course.name + "\t" + progfraction + "\t" + course.courseincomedict[bestcode] + "\t" + course.coursecostdict[bestcode]
                    //    + "\t" + (course.courseincomedict[bestcode] - course.coursecostdict[bestcode]));
                    double semsum = 0;
                    for (int i=0;i<=6; i++)
                    {
                        if (pc.coursedict.ContainsKey(i) && pc.coursedict[i].ContainsKey(bestcode))
                            semsum += pc.coursedict[i][bestcode];
                    }
                    if (semsum > 0)
                    {
                        for (int i = 0; i <= 6; i++)
                        {
                            if (pc.coursedict.ContainsKey(i) && pc.coursedict[i].ContainsKey(bestcode))
                                semcost[i] += courseprogcost*pc.coursedict[i][bestcode]/semsum;
                        }
                    }
                }
                double tb = sumincome - sumcost;
                double costplusoh = 1.6 * sumcost;
                double profit = sumincome - costplusoh;
                sb.Append("\t" + sumhst + "\t" + sumincome + "\t" + sumcost + "\t" + tb+"\t"+costplusoh+"\t"+profit);
                for (int i = 1; i <= 6; i++)
                {
                    if (Math.Abs(semcost[i]) > 1000)
                        sb.Append("\t" + semcost[i]);
                    else
                        sb.Append("\t");
                }
                memo(sb.ToString());

            }
        }

        private void read_kurskostnad()
        {
            openFileDialog1.Title = "Select retendo_ladok file";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fn = openFileDialog1.FileName;
                memo("Reading program/course-income data from " + fn);

                read_retendo_ladok(fn);
            }
        }

        private void read_retendo_ladok(string fn)
        {
            List<string> courses = new List<string>();

            memo("Reading " + fn);

            int nfoundcode = 0;
            int nfoundname = 0;
            int nnfound = 0;
            int nprognotfound = 0;
            using (StreamReader sr = new StreamReader(fn))
            {
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
                    if (words.Length < 9)
                        continue;
                    string code = words[1];
                    if (string.IsNullOrEmpty(code))
                        continue;
                    string coursename = words[0].Replace(code, "").Trim();
                    double hp = -1;
                    if (coursename.Contains("hp"))
                    {
                        var hpresult = util.extract_hp(coursename);
                        coursename = hpresult.Item1;
                        hp = hpresult.Item2;
                    }
                    double hst = util.tryconvertdouble(words[2]);
                    double hpr = util.tryconvertdouble(words[3]);
                    double income = util.tryconvertdouble(words[7]);
                    double cost = util.tryconvertdouble(words[8]);
                    if (words.Length > 21)
                    {
                        income = util.tryconvertdouble(words[18]);
                        cost = util.tryconvertdouble(words[9]);
                        //cost = util.tryconvertdouble(words[21]); With OH!
                    }
                    programclass course = findcourse(code);
                    if (course == null)
                    {
                        course = findcourse(coursename);
                        //memo(coursename + " sought by name");
                        if (course != null)
                        {
                            //if (hp == course.hp || course.hp < 0 || hp < 0)
                            {
                                nfoundname++;
                                course.coursecodelist.Add(code);
                                course.hp = hp;
                                fkcodedict.Add(code, course);
                            }
                            //else
                            //    course = null;
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
                        course.hp = hp;
                        //course.subject = words[0];
                        //course.sector = words[1];
                        course.coursecodelist.Add(code);
                        course.subjectcode = getsubjectcode(code);
                        course.homeinst = (shortinstdict[subjinstdict[course.subjectcode]]);
                        var c2 = (from c in fkdict.Values
                                  where c.subjectcode == course.subjectcode
                                  select c).FirstOrDefault();
                        if (c2 != null)
                            course.studentpengarea = c2.studentpengarea;
                        fkdict.Add(coursename, course);
                        fkcodedict.Add(code, course);
                        //continue;
                    }
                    string bestcode = course.bestcode();
                    //public Dictionary<string, double> coursehstdict = new Dictionary<string, double>();
                    //public Dictionary<string, double> courseincomedict = new Dictionary<string, double>();
                    //public Dictionary<string, double> coursecostdict = new Dictionary<string, double>();
                    if (!course.coursehstdict.ContainsKey(bestcode))
                    {
                        course.coursehstdict.Add(bestcode, hst);
                        course.courseincomedict.Add(bestcode, income);
                        course.coursecostdict.Add(bestcode, cost);
                    }
                    else
                    {
                        course.coursehstdict[bestcode] += hst;
                        course.courseincomedict[bestcode] += income;
                        course.coursecostdict[bestcode] += cost;
                    }

                }
                memo("# lines " + nline);
                //memo("# progs = " + progs.Count);
                memo("#courses found by code: " + nfoundcode);
                memo("#courses found by name: " + nfoundname);
                memo("#courses not found: " + nnfound);
                memo("#programs not found: " + nprognotfound);

            }
        }

        private void read_applicants()
        {
            openFileDialog1.Title = "Select anmälningar file";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fn = openFileDialog1.FileName;
                memo("Reading applicant data from " + fn);

                using (StreamReader sr = new StreamReader(fn))
                {
                    sr.ReadLine(); //throw away headerlines
                    string year = sr.ReadLine().Substring(7,2);
                    sr.ReadLine();

                    string hline = sr.ReadLine();
                    string[] hwords = hline.Split('\t');

                    int nline = 0;
                    int ngood = 0;

                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        string[] words = line.Split('\t');
                        nline++;
                        if (words.Length < 4)
                            continue;
                        string code = words[1];
                        if (string.IsNullOrEmpty(code))
                            continue;
                        string coursename = words[0].Replace(code, "").Trim();
                        double hp = -1;
                        if (coursename.Contains("hp"))
                        {
                            var hpresult = util.extract_hp(coursename);
                            coursename = hpresult.Item1;
                            hp = hpresult.Item2;
                        }
                        programclass course = findcourse(code);
                        if (course == null)
                        {
                            course = findcourse(coursename);
                        }
                        if (course == null)
                            continue;

                        string sem = words[0].Contains("(V") ? "VT" : "HT";
                        sem += year;
                        programbatchclass pb = course.getbatch(sem);
                        if (pb == null)
                            continue;
                        pb.appldict = new dictclass(hwords, words);
                        ngood++;
                    }
                    memo("Lines: " + nline);
                    memo("Good:  " + ngood);
                }

            }

        }

        private void fill_fk_progstud(string sem)
        {
            foreach (programclass pc in origprogramdict.Values)
            {
                foreach (programbatchclass pb in pc.batchlist)
                {
                    double nstud = pb.getstud(sem);
                    int tx = util.semestercount(pb.batchstart, sem);
                    if (pc.coursedict.ContainsKey(tx))
                    {
                        foreach (string code in pc.coursedict[tx].Keys)
                        {
                            programclass course = findcourse(code);
                            if (code == null)
                                continue;
                            var pbx = (from c in course.batchlist where c.batchstart == sem select c).FirstOrDefault();
                            if (pbx == null)
                                continue;
                            pbx.progstud += nstud * pc.coursedict[tx][code];
                        }
                    }
                }
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            string yearstring = "23";
            var q = from c in origprogramdict.Values where c.courseincomedict.Count() > 0 select c;
            if (q.Count() == 0)
            {
                //read_lokal_ers_programkurser();
                read_kurskostnad();
                read_applicants();
                fill_fk_progstud("VT"+yearstring);
                fill_fk_progstud("HT"+yearstring);
            }

            int ncell1 = 4;
            int ncelltot = 4;
            int[,] coursecount = new int[ncell1, ncelltot];
            double[,] withprofit = new double[ncell1, ncelltot];
            double[,] profitsum = new double[ncell1, ncelltot];
            double[,] incomesum = new double[ncell1, ncelltot];
            double[,] hstsum = new double[ncell1, ncelltot];



            int low1 = 10;
            int pitch1 = 10;
            int lowtot = 50;
            int pitchtot = 25;

            string s1 = "Antal behöriga sökande 1a-hand";
            string stot = "Antal sökande totalt";

            foreach (programclass course in fkdict.Values)
            {
                if (course.courseincomedict.Count == 0)
                    continue;
                //if (course.homeinst != "Institutionen för språk, litteratur och lärande")
                //    continue;
                string bestcode = course.bestcode();
                double income = course.courseincomedict[bestcode];
                double profit = income - 1.6 * course.coursecostdict[bestcode];
                foreach (programbatchclass pb in course.batchlist)
                {
                    if (!pb.batchstart.Contains(yearstring))
                        continue;
                    if (pb.appldict == null)
                        continue;
                    if (pb.progstud > 0.5 * pb.getstud(1))
                        continue;
                    int n1 = pb.appldict.Getint(s1);
                    int ntot = pb.appldict.Getint(stot);
                    if (n1 < 20 && ntot < 100)
                        memo(bestcode+"\t"+course.name+"\t"+n1+"\t"+ntot);
                    int i1 = 0;
                    if (n1 > low1)
                    {
                        i1 = 1 + (n1 - low1) / pitch1;
                        if (i1 >= ncell1)
                            i1 = ncell1 - 1;
                    }
                    int itot = 0;
                    if (ntot > lowtot)
                    {
                        itot = 1 + (ntot - lowtot) / pitchtot;
                        if (itot >= ncelltot)
                            itot = ncelltot - 1;
                    }
                    coursecount[i1, itot]++;
                    if (profit>0)
                        withprofit[i1,itot]++;
                    if (course.hp > 0)
                        hstsum[i1, itot] += pb.getstud(1) * course.hp / 60;
                    incomesum[i1, itot] += income;
                    profitsum[i1, itot] += profit;
                    break;
                }

            }
            StringBuilder sb1 = new StringBuilder();
            int k = low1;
            do
            {
                sb1.Append("\t" + k);
                k += pitch1;
            }
            while (k <= low1 + ncell1 * pitch1);
            memo(sb1.ToString());

            int ktot = lowtot;
            int j = 0;
            StringBuilder sb = new StringBuilder();
            do
            {
                sb = new StringBuilder(ktot.ToString());
                for (int i=0; i<ncell1; i++)
                {
                    if (coursecount[i, j] > 0)
                        sb.Append("\t" + (withprofit[i, j] / coursecount[i, j]).ToString());
                    else
                        sb.Append("\t");
                }
                memo(sb.ToString());
                ktot += pitchtot;
                j += 1;
            }
            while (j < ncelltot);

            memo("");
            memo("");
            memo(sb1.ToString());

            sb = new StringBuilder();
            ktot = lowtot;
            j = 0;
            do
            {
                sb = new StringBuilder(ktot.ToString());
                for (int i = 0; i < ncell1; i++)
                {
                    if (coursecount[i, j] > 0)
                        sb.Append("\t" + (profitsum[i, j] / coursecount[i, j]).ToString());
                    else
                        sb.Append("\t");
                }
                memo(sb.ToString());
                ktot += pitchtot;
                j += 1;
            }
            while (j < ncelltot);

            memo("");
            memo("Antal kurser per cell");
            memo(sb1.ToString());

            sb = new StringBuilder();
            ktot = lowtot;
            j = 0;
            do
            {
                sb = new StringBuilder(ktot.ToString());
                for (int i = 0; i < ncell1; i++)
                {
                    if (coursecount[i, j] > 0)
                        sb.Append("\t" + coursecount[i, j].ToString());
                    else
                        sb.Append("\t");
                }
                memo(sb.ToString());
                ktot += pitchtot;
                j += 1;
            }
            while (j < ncelltot);

            memo("");
            memo("Total intäkt per cell");
            memo(sb1.ToString());

            sb = new StringBuilder();
            ktot = lowtot;
            j = 0;
            do
            {
                sb = new StringBuilder(ktot.ToString());
                for (int i = 0; i < ncell1; i++)
                {
                    if (coursecount[i, j] > 0)
                        sb.Append("\t" + incomesum[i, j].ToString());
                    else
                        sb.Append("\t");
                }
                memo(sb.ToString());
                ktot += pitchtot;
                j += 1;
            }
            while (j < ncelltot);

            memo("");
            memo("HST per cell");
            memo(sb1.ToString());

            sb = new StringBuilder();
            ktot = lowtot;
            j = 0;
            do
            {
                sb = new StringBuilder(ktot.ToString());
                for (int i = 0; i < ncell1; i++)
                {
                    if (coursecount[i, j] > 0)
                        sb.Append("\t" + hstsum[i, j].ToString());
                    else
                        sb.Append("\t");
                }
                memo(sb.ToString());
                ktot += pitchtot;
                j += 1;
            }
            while (j < ncelltot);

        }

        private void ProgInstbutton_Click(object sender, EventArgs e)
        {
            foreach (programclass pc in origprogramdict.Values)
            {
                Dictionary<string, double> studdict = new Dictionary<string, double>();
                double totstud = 0;
                studdict.Add(pc.homeinst, 0);
                foreach (int sem in pc.coursedict.Keys)
                {
                    foreach (string code in pc.coursedict[sem].Keys)
                    {
                        totstud += pc.coursedict[sem][code]*fkcodedict[code].hp;
                        if (fkcodedict[code].homeinst == pc.homeinst)
                        {
                            studdict[pc.homeinst] += pc.coursedict[sem][code] * fkcodedict[code].hp;
                        }
                        else
                        {
                            string cc = code + " " + fkcodedict[code].name;
                            if (studdict.ContainsKey(cc))
                                studdict[cc] += pc.coursedict[sem][code] * fkcodedict[code].hp;
                            else
                                studdict.Add(cc,pc.coursedict[sem][code] * fkcodedict[code].hp);
                            cc = fkcodedict[code].homeinst;
                            if (studdict.ContainsKey(cc))
                                studdict[cc] += pc.coursedict[sem][code] * fkcodedict[code].hp;
                            else
                                studdict.Add(cc, pc.coursedict[sem][code] * fkcodedict[code].hp);
                        }
                    }
                }
                if (totstud > 0)
                {
                    memo(pc.name+"\t"+instshortdict[pc.homeinst]);
                    foreach (string inst in instshortdict.Keys)
                    {
                        if (studdict.ContainsKey(inst) && studdict[inst] > 0)
                        {
                            memo("\t" + instshortdict[inst] + "\t" + (100 * studdict[inst] / totstud) + " %");
                        }
                    }
                    foreach (string cc in studdict.Keys)
                    {
                        if (!instshortdict.ContainsKey(cc) && studdict[cc] > 0)
                        {
                            memo("\t" + cc + "\t" + (100 * studdict[cc] / totstud) + " %");
                        }
                    }
                    memo("");
                }
            }
        }

        public static Dictionary<string, string> trianglecolor = new Dictionary<string, string>()
        {
            {"Ekonomi och arbetsliv","#FF0000" },
            {"Samhälle","#FF5500" },
            {"Medier","#FF0055" },
            {"Lärarutb","#FF8800" },
            //{"IKS lärarutb","#FF8800" },
            {"IKS resurs","#AA8800" },
            {"Vård","#00FF00" },
            {"Socionom","#55FF00" },
            {"Idrott och hälsa","#00FF55" },
            //{"IHOV lärarutb","#00FF88" },
            {"Energi och bygg","#0000FF" },
            {"Industri","#0000AA" },
            {"Data och IT","#0055FF" },
            {"IIT resurs","#0055AA" },
            {"Språk","#FFFF00" },
        };

        public static Dictionary<string, string> progcolor = new Dictionary<string, string>()
        {
            {"Ekonomi och arbetsliv","#FF5555" },
            {"Samhälle","#FFAA55" },
            {"Medier","#FF55AA" },
            {"Lärarutb","#FFCC55" },
            //{"IKS lärarutb","#FFCC55" },
            {"IKS resurs","#BB9955" },
            {"Vård","#55FF55" },
            {"Socionom","#99FF55" },
            {"Idrott och hälsa","#55FF99" },
            //{"IHOV lärarutb","#55FFAA" },
            {"Energi och bygg","#5555FF" },
            {"Industri","#3333AA" },
            {"Data och IT","#5599FF" },
            {"IIT resurs","#5599BB" },
            {"Språk","#FFFF55" },
        };

        private Dictionary<string, string> triangleindex = new Dictionary<string, string>()
        {

{"LP","Medier"},
{"SR","Vård"},
{"ST","Data och IT"},
{"MI","Data och IT"},
{"IK","Data och IT"},
{"DT","Data och IT"},
{"NA","Ekonomi och arbetsliv"},
{"SQ","Energi och bygg"},
{"BY","Energi och bygg"},
{"IE","Industri"},
{"MA","IIT resurs"},
{"FI","IKS resurs"},
{"FÖ","Ekonomi och arbetsliv"},
{"EU","Ekonomi och arbetsliv"},
{"RV","Ekonomi och arbetsliv"},
{"MT","Industri"},
{"MD","IIT resurs"},
{"EG","Energi och bygg"},
{"PA","Ekonomi och arbetsliv"},
{"ET","Industri"},
{"FY","IIT resurs"},
{"TR","Ekonomi och arbetsliv"},
{"IH","Idrott och hälsa"},
{"SO","Samhälle"},
{"SK","Samhälle"},
{"AB","Ekonomi och arbetsliv"},
{"MÖ","Energi och bygg"},
{"RY","Språk"},
{"BQ","Medier"},
{"BP","Medier"},
{"EN","Språk"},
{"PG","IKS lärarutb"},
{"SV","Språk"},
{"NV","IHOV lärarutb"},
{"GT","Data och IT"},
{"HI","IKS resurs"},
{"SS","Språk"},
{"KG","Ekonomi och arbetsliv"},
{"AU","Medier"},
{"MC","Vård"},
{"MP","Industri"},
{"RK","IKS resurs"},
{"SA","Socionom"},
{"VÅ","Vård"},
{"PS","Socionom"},
{"FR","Språk"},
{"KT","Industri"},
{"KE","IHOV lärarutb"},
{"TY","Språk"},
{"AR","Språk"},
{"SP","Språk"},
{"KI","Språk"},
{"IT","Språk"},
{"LI","Språk"},
{"GG","IKS resurs"},
{"JP","Språk"},
{"PR","Språk"},
{"Audiovisuell produktion","Medier"},
{"Barnmorskeprogrammet","Vård"},
{"Magisterpgm Business Intelligence","Data och IT"},
{"Bygg- och samhällsplanerarprogrammet","Energi och bygg"},
{"Byggarbetsledarprogrammet","Energi och bygg"},
{"Byggingenjör","Energi och bygg"},
{"Byggteknik - Produktion och förvaltning","Energi och bygg"},
{"Masterprogram i Data Science","Data och IT"},
{"Detaljhandelsprogrammet 180hp","Ekonomi och arbetsliv"},
{"Detaljhandelprogrammet 120hp","Ekonomi och arbetsliv"},
{"Digitalbrott och eSäkerhet","Data och IT"},
{"Ekonomprogrammet","Ekonomi och arbetsliv"},
{"Högskoleingenjör hållbara energisystem","Energi och bygg"},
{"Energiteknikerprogrammet","Energi och bygg"},
{"Entreprenöriellt företagande","Ekonomi och arbetsliv"},
{"Film- och TV produktion","Medier"},
{"Filmdesign för reklam och information","Medier"},
{"Förskollärarprogrammet","IKS lärarutb"},
{"Global sexuell och reproduktiv hälsa","Vård"},
{"Grafisk design","IKS lärarutb"},
{"Grundlärare Förskoleklass - årskurs 1-3","IKS lärarutb"},
{"Grundlärare Grundskolans årskurs 4-6","IKS lärarutb"},
{"Grundlärare årskurs 4-6 arbetsintegrerad","IKS lärarutb"},
{"Grundlärare årskurs F-3 arbetsintegrerad","IKS lärarutb"},
{"Idrottstränarpgm. m. inr mot prestation","Idrott och hälsa"},
{"Idrottstränarpgm m. hälsoinriktning","Idrott och hälsa"},
{"Industriell ekonomi - Högskoleexamen","Industri"},
{"Industriell ekonomi","Industri"},
{"Informationsdesign - kandidatprogram","Data och IT"},
{"International Tourism Management","Ekonomi och arbetsliv"},
{"IT säkerhet och mjukvarutestning","Data och IT"},
{"Manus för film och TV","Medier"},
{"Kompl lärarutb år 7-9","IKS lärarutb"},
{"Kompl lärarutb gymnasiet","IKS lärarutb"},
{"Kompl lärarutb yrkes år 7-9","IKS lärarutb"},
{"Ljud- och musikproduktionsprogrammet","Medier"},
{"Lärarprogrammet","IKS lärarutb"},
{"Magisterpgm i Audiovisuella Studier","Medier"},
{"Magisterpgm i destinationsutveckling","Ekonomi och arbetsliv"},
{"Magisterpgm i Energieffektivt byggande","Energi och bygg"},
{"Magisterpgm i engelska - lingvistik","Språk"},
{"Magisterpgm i engelska - eng litteratur","Språk"},
{"Magisterpgm i fysioterapi","Vård"},
{"Magisterpgm i Företagsekonomi","Ekonomi och arbetsliv"},
{"Magisterpgm Maskin/Materialteknik","Industri"},
{"Magisterpgm i Nationalekonomi","Ekonomi och arbetsliv"},
{"Magisterpgm  i religionsvetenskap","IKS resurs"},
{"Magisterpgm i Pedagogiskt arbete","IKS lärarutb"},
{"Magisterpgm i solenergiteknik","Energi och bygg"},
{"Magisterpgm Svenska som andraspråk","Språk"},
{"Magisterpgm Demokrati, medborgarskap och förändring","Samhälle"},
{"Högskoleingenjör - Maskinteknik","Industri"},
{"Master Business Intelligence","Data och IT"},
{"Masterprogram i Materialteknik och Produktutveckling","Industri"},
{"Masterprogram i solenergiteknik","Energi och bygg"},
{"Musik- och ljuddesign","Medier"},
{"Personal- och arbetslivsprogram","Ekonomi och arbetsliv"},
{"Samhällsvetarprogrammet","Samhälle"},
{"Sjuksköterskeprogrammet","Vård"},
{"Socionomprogrammet","Socionom"},
{"Spec.sjuksköterska - vård av äldre","Vård"},
{"Spec.sjuksköterska - demensvård","Vård"},
{"Spec.sjuksköterska - distriktsköterska","Vård"},
{"Sport Management","Ekonomi och arbetsliv"},
{"Systemvetenskapliga programmet","Data och IT"},
{"Teknisk/Naturvetenskaplig bastermin","IIT resurs"},
{"Tekniskt basår","IIT resurs"},
{"Tränarprogrammet för hälsa och idrott","Idrott och hälsa"},
{"Digitala tjänster 120 hp","Data och IT"},
{"Vidareutbildning av lärare","IKS lärarutb"},
{"Yrkeslärare","IKS lärarutb"},
{"Ämneslärare - Grundskolans årskurs 7-9","IKS lärarutb"},
{"Ämneslärare Gymnasieskolan","IKS lärarutb"},
{"Upplevelseproduktion - friluftsliv, idrott och hälsa","Ekonomi och arbetsliv"},
{"Produktionstekniker 120 hp","Industri"},
{"Ekonomprogrammet - Business Management","Ekonomi och arbetsliv"},
{"Masterprogram i interkulturella litteraturstudier","Språk"},
{"Filmproduktion 180 hp","Medier"},
{"Manus och producentskap 180 hp","Medier"},
{"Medieproduktion för TV och sociala medier 180 hp","Medier"},
{"Musik- och ljudproduktion 180 hp","Medier"},
{"Statsvetenskap A - distans","Samhälle"},
{"Statsvetenskap B - Distans","Samhälle"},
{"Statsvetenskap C - Distans","Samhälle"},
{"Spanska I","Språk"},
{"Spanska II","Språk"},
{"Svenska som andraspråk I","Språk"},
{"Svenska som andraspråk II","Språk"},
{"Portugisiska Grundläggande kurs I och II, 30hp","Språk"},
{"Portugisiska I","Språk"},
{"Portugisiska II","Språk"},
{"Kinesiska I - Distans","Språk"},
{"Kinesiska I - Distans, Kväll","Språk"},
{"Kinesiska II - Distans","Språk"},
{"Kinesiska II - Distans, Kväll","Språk"},
{"Kinesiska III - Distans","Språk"},
{"Kinesiska IV - Distans","Språk"},
{"Franska I","Språk"},
{"Franska II","Språk"},
{"Franska IIIa - Fransk litteraturvetenskap","Språk"},
{"Franska IIIb - Fransk grammatik och lingvistik","Språk"},
{"Japanska I - Distans","Språk"},
{"Japanska II - Distans","Språk"},
{"Engelska II","Språk"},
{"Tränare i styrkelyft - Kurspaket","Idrott och hälsa"},
{"Judotränare  - Kurspaket","Idrott och hälsa"},
{"Tyska I","Språk"},
{"Tyska I, 50%","Språk"},
{"Tyska II, 50%","Språk"},
{"Tyska II","Språk"},
{"Tyska III","Språk"},
{"Tyska: Introduktionspaket I i Tyska","Språk"},
{"Tyska: Introduktionspaket II i Tyska, 25%","Språk"},
{"Kinesiska I: Vägen till Kina - Distans","Språk"},
{"Tränare i klättring - Kurspaket","Idrott och hälsa"},
{"Statsvetenskap III","Samhälle"},
{"Energy for exchange students 2","Energi och bygg"},
{"Japanska III - Distans","Språk"},
{"Engelska I","Språk"},
{"Statsvetenskap II - Distans","Samhälle"}


        };



        private void progsubjbutton_Click(object sender, EventArgs e)
        {
            Dictionary<string, Dictionary<string, double>> progsubjdict = new Dictionary<string, Dictionary<string, double>>();
            List<string> subjectlist = new List<string>();

            //Gather course data by subject:
            foreach (programclass pc in origprogramdict.Values)
            {
                if (pc.coursedict.Count == 0)
                    continue;
                progsubjdict.Add(pc.name, new Dictionary<string, double>());
                foreach (int i in pc.coursedict.Keys)
                {
                    foreach (string ccode in pc.coursedict[i].Keys)
                    {
                        string subj = fkcodedict[ccode].subjectcode;
                        if (!subjectlist.Contains(subj))
                            subjectlist.Add(subj);
                        if (!progsubjdict[pc.name].ContainsKey(subj))
                            progsubjdict[pc.name].Add(subj, pc.coursedict[i][ccode]);
                        else
                            progsubjdict[pc.name][subj] += pc.coursedict[i][ccode];
                    }
                }
            }

            //Make Sankey?

            //Normalize:
            foreach (var dd in progsubjdict.Values)
            {
                double sum = dd.Values.Sum();
                if (sum == 0)
                    continue;
                foreach (string ccode in dd.Keys.ToList())
                {
                    dd[ccode] *= 100 / sum; // *100 to make percentage values
                }
            }

            //Make Gephi input:

            string fnnode = util.unusedfn(folder + @"\progsubjnodes.csv");
            string fnedge = util.unusedfn(folder + @"\progsubjedges.csv");
            string fnsubjnode = util.unusedfn(folder + @"\subjnodes.csv");
            string fnsubjedge = util.unusedfn(folder + @"\subjedges.csv");


            Dictionary<string, int> nodedict = new Dictionary<string, int>();
            using (StreamWriter sw = new StreamWriter(fnnode))
            using (StreamWriter sws = new StreamWriter(fnsubjnode))
            {
                memo("Writing " + fnnode);
                memo("Writing " + fnsubjnode);
                int inode = 0;
                sw.WriteLine("Id;Label;Color");
                sws.WriteLine("Id;Label;Color");
                foreach (string subj in subjectlist)
                {
                    inode++;
                    nodedict.Add(subj, inode);
                    sw.WriteLine(inode + ";" + subj + ";"+trianglecolor[triangleindex[subj]]);
                    sws.WriteLine(inode + ";" + subj + ";" + trianglecolor[triangleindex[subj]]);
                }
                foreach (string progname in progsubjdict.Keys)
                {
                    inode++;
                    nodedict.Add(progname, inode);
                    sw.WriteLine(inode + ";" + progname + ";" + trianglecolor[triangleindex[progname]]);
                }
            }

            using (StreamWriter sw = new StreamWriter(fnedge))
            {
                memo("Writing " + fnedge);
                sw.WriteLine("Source;Target;Weight");
                foreach (string progname in progsubjdict.Keys)
                {
                    int iprog = nodedict[progname];
                    foreach (string subj in progsubjdict[progname].Keys)
                    {
                        int isubj = nodedict[subj];
                        int w = (int)Math.Round(progsubjdict[progname][subj]);
                        if (w > 0)
                            sw.WriteLine(iprog + ";" + isubj + ";" + w);
                    }
                }
            }

            using (StreamWriter sw = new StreamWriter(fnsubjedge))
            {
                memo("Writing " + fnsubjedge);
                sw.WriteLine("Source;Target;Weight");
                foreach (string progname in progsubjdict.Keys)
                {
                    int iprog = nodedict[progname];
                    foreach (string subj in progsubjdict[progname].Keys)
                    {
                        int isubj = nodedict[subj];
                        foreach (string subj2 in progsubjdict[progname].Keys)
                        {
                            if (subj == subj2)
                                continue;
                            int isubj2 = nodedict[subj2];
                            if (isubj2 < isubj)
                                continue;
                            int w = (int)Math.Round(progsubjdict[progname][subj]*progsubjdict[progname][subj2]);
                            if (w > 0)
                                sw.WriteLine(isubj + ";" + isubj2 + ";" + w);
                        }

                    }
                }
            }

        }

        public static Dictionary<string, string> subjectriangle = new Dictionary<string, string>();
        public static Dictionary<string, string> subjcodetriangle = new Dictionary<string, string>();
        public static Dictionary<string, string> progtriangle = new Dictionary<string, string>();

        private void fill_triangledicts()
        {
            if (subjectriangle.Count > 0)
                return;

            subjectriangle.Add("Bild", "Medier");
            subjectriangle.Add("Pedagogiskt arbete", "Lärarutb");
            subjectriangle.Add("Pedagogik", "Lärarutb");
            subjectriangle.Add("Arbetsvetenskap", "Ekonomi och arbetsliv");
            subjectriangle.Add("Sociologi", "Samhälle");
            subjectriangle.Add("Turismvetenskap", "Ekonomi och arbetsliv");
            subjectriangle.Add("Kulturgeografi", "Ekonomi och arbetsliv");
            subjectriangle.Add("Datateknik", "Data och IT");
            subjectriangle.Add("Informatik", "Data och IT");
            subjectriangle.Add("Mikrodataanalys", "Data och IT");
            subjectriangle.Add("Statistik", "Data och IT");
            subjectriangle.Add("Byggteknik", "Energi och bygg");
            subjectriangle.Add("Energiteknik", "Energi och bygg");
            subjectriangle.Add("Samhällsbyggnadsteknik", "Energi och bygg");
            subjectriangle.Add("Filosofi", "IKS resurs");
            subjectriangle.Add("Historia", "IKS resurs");
            subjectriangle.Add("Religion", "IKS resurs");
            subjectriangle.Add("Religionsvetenskap", "IKS resurs");
            subjectriangle.Add("Statsvetenskap", "Samhälle");
            subjectriangle.Add("Entreprenörsskap och innovationsteknik", "Ekonomi och arbetsliv");
            subjectriangle.Add("Entreprenörskap och innovationsteknik", "Ekonomi och arbetsliv");
            subjectriangle.Add("Fysik", "IIT resurs");
            subjectriangle.Add("Industriell ekonomi", "Industri");
            subjectriangle.Add("Maskinteknik", "Industri");
            subjectriangle.Add("Elektroteknik", "Industri");
            subjectriangle.Add("Matematik", "IIT resurs");
            subjectriangle.Add("Materialteknik", "Industri");
            subjectriangle.Add("Bildproduktion", "Medier");
            subjectriangle.Add("Ljud- och musikproduktion", "Medier");
            subjectriangle.Add("Företagsekonomi", "Ekonomi och arbetsliv");
            subjectriangle.Add("Nationalekonomi", "Ekonomi och arbetsliv");
            subjectriangle.Add("Rättsvetenskap", "Ekonomi och arbetsliv");
            subjectriangle.Add("Omvårdnad", "Vård");
            subjectriangle.Add("Arabiska", "Språk");
            subjectriangle.Add("Franska", "Språk");
            subjectriangle.Add("Italienska", "Språk");
            subjectriangle.Add("Kinesiska", "Språk");
            subjectriangle.Add("Portugisiska", "Språk");
            subjectriangle.Add("Spanska", "Språk");
            subjectriangle.Add("Engelska", "Språk");
            subjectriangle.Add("Japanska", "Språk");
            subjectriangle.Add("Ryska", "Språk");
            subjectriangle.Add("Tyska", "Språk");
            subjectriangle.Add("Svenska som andraspråk", "Språk");
            subjectriangle.Add("Svenska språket", "Språk");
            subjectriangle.Add("Svenska", "Språk");
            subjectriangle.Add("Matematikdidaktik", "Lärarutb");
            subjectriangle.Add("Naturvetenskap", "Lärarutb");
            subjectriangle.Add("Idrott och hälsa", "Idrott och hälsa");
            subjectriangle.Add("Idrotts- och hälsovetenskap", "Idrott och hälsa");
            subjectriangle.Add("Medicinsk vetenskap", "Vård");
            subjectriangle.Add("Sexuell, reproduktiv och perinatal hälsa", "Vård");
            subjectriangle.Add("Socialt arbete", "Socionom");

            progtriangle.Add("Digitala tjänster - program för högskoleexamen 120 hp", "Data och IT");
            progtriangle.Add("Grafisk design och webbutveckling 180 hp", "Data och IT");
            progtriangle.Add("IT säkerhet och mjukvarutestning 180 hp", "Data och IT");
            progtriangle.Add("Masterprogram i Data Science 120 hp", "Data och IT");
            progtriangle.Add("Systemvetenskapliga programmet 180 hp", "Data och IT");
            progtriangle.Add("Ekonomprogrammet - Business Management 180 hp", "Ekonomi och arbetsliv");
            progtriangle.Add("Ekonomprogrammet 180 hp", "Ekonomi och arbetsliv");
            progtriangle.Add("Entreprenöriellt företagande 120 hp", "Ekonomi och arbetsliv");
            progtriangle.Add("International Tourism Management 180 hp", "Ekonomi och arbetsliv");
            progtriangle.Add("Magisterprogram Business Intelligence 60 hp", "Ekonomi och arbetsliv");
            progtriangle.Add("Magisterprogram i destinationsutveckling 60 hp", "Ekonomi och arbetsliv");
            progtriangle.Add("Magisterprogram i Företagsekonomi 60 hp", "Ekonomi och arbetsliv");
            progtriangle.Add("Magisterprogram i Nationalekonomi 60 hp", "Ekonomi och arbetsliv");
            progtriangle.Add("Personal- och arbetslivsprogram 180 hp", "Ekonomi och arbetsliv");
            progtriangle.Add("Sport Management 180 hp", "Ekonomi och arbetsliv");
            progtriangle.Add("Bygg- och samhällsplanerarprogrammet 180 hp", "Energi och bygg");
            progtriangle.Add("Byggingenjör 180 hp", "Energi och bygg");
            progtriangle.Add("Byggteknik - Produktion och förvaltning 120 hp", "Energi och bygg");
            progtriangle.Add("Energiteknik - Högskoleingenjör 180 hp", "Energi och bygg");
            progtriangle.Add("Magisterprogram i Energieffektivt byggande 60 hp", "Energi och bygg");
            progtriangle.Add("Magisterprogram i solenergiteknik 60 hp", "Energi och bygg");
            progtriangle.Add("Masterprogram i solenergiteknik 120 hp", "Energi och bygg");
            progtriangle.Add("Tränarprogrammet för hälsa och idrott 180 hp", "Idrott och hälsa");
            progtriangle.Add("Upplevelseproduktion - friluftsliv, idrott och hälsa 180 hp", "Idrott och hälsa");
            progtriangle.Add("Tekniskt basår 30 hp", "IIT resurs");
            progtriangle.Add("Tekniskt basår 60 hp", "IIT resurs");
            progtriangle.Add("Ämneslärare - Grundskolans årskurs 7-9 240 hp", "Lärarutb");
            progtriangle.Add("Ämneslärare - Grundskolans årskurs 7-9 270 hp", "Lärarutb");
            progtriangle.Add("Ämneslärare Gymnasieskolan 330 hp", "Lärarutb");
            progtriangle.Add("Grundlärare årskurs 4-6 arbetsintegrerad 240 hp", "Lärarutb");
            progtriangle.Add("Grundlärare årskurs F-3 arbetsintegrerad 240 hp", "Lärarutb");
            progtriangle.Add("Grundlärare Förskoleklass - årskurs 1-3 240 hp", "Lärarutb");
            progtriangle.Add("Grundlärare Grundskolans årskurs 4-6 240 hp", "Lärarutb");
            progtriangle.Add("Kompl lärarutb år 7-9 90 hp", "Lärarutb");
            progtriangle.Add("Kompl lärarutb gymnasiet 90 hp", "Lärarutb");
            progtriangle.Add("Magisterprogram i Pedagogiskt arbete 60 hp", "Lärarutb");
            progtriangle.Add("Vidareutbildning av lärare 330 hp", "Lärarutb");
            progtriangle.Add("Yrkeslärare 90 hp", "Lärarutb");
            progtriangle.Add("Industriell ekonomi - Högskoleexamen 120 hp", "Industri");
            progtriangle.Add("Produktionsteknikerprogrammet 120 hp", "Industri");
            progtriangle.Add("Audiovisuell produktion 180 hp", "Medier");
            progtriangle.Add("Film- och TV produktion 180 hp", "Medier");
            progtriangle.Add("Filmproduktion för reklam och information - kandidatprogram 180 hp", "Medier");
            progtriangle.Add("Ljud- och musikproduktionsprogrammet 180 hp", "Medier");
            progtriangle.Add("Magisterprogram i Audiovisuella Studier 60 hp", "Medier");
            progtriangle.Add("Manus för film och TV 180 hp", "Medier");
            progtriangle.Add("Musik- och ljuddesign 120 hp", "Medier");
            progtriangle.Add("Magisterprogram Demokrati, medborgarskap och förändring 60 hp", "Samhälle");
            progtriangle.Add("Samhällsvetarprogrammet 180 hp", "Samhälle");
            progtriangle.Add("Socionomprogrammet 210 hp", "Socionom");
            progtriangle.Add("Magisterpgm Svenska som andraspråk 60 hp", "Språk");
            progtriangle.Add("Magisterprogram i engelska - engelskspråkig litteratur 60 hp", "Språk");
            progtriangle.Add("Magisterprogram i engelska - lingvistik 60 hp", "Språk");
            progtriangle.Add("Masterprogram i interkulturella litteraturstudier 120 hp", "Språk");
            progtriangle.Add("Barnmorskeprogrammet 90 hp", "Vård");
            progtriangle.Add("Global sexuell och reproduktiv hälsa 60 hp", "Vård");
            progtriangle.Add("Magisterprogram i fysioterapi 60 hp", "Vård");
            progtriangle.Add("Sjuksköterskeprogrammet 180 hp", "Vård");
            progtriangle.Add("Spec.sjuksköterska - demensvård 60 hp", "Vård");
            progtriangle.Add("Spec.sjuksköterska - distriktsköterska 75 hp", "Vård");
            progtriangle.Add("Spec.sjuksköterska - vård av äldre 60 hp", "Vård");
            progtriangle.Add("Ämneslärare - Grundskolans årskurs 7-9", "Lärarutb");
            progtriangle.Add("Ämneslärare Gymnasieskolan", "Lärarutb");
            progtriangle.Add("Audiovisuell produktion", "Medier");
            progtriangle.Add("Barnmorskeprogrammet", "Vård");
            progtriangle.Add("Bygg- och samhällsplanerarprogrammet", "Energi och bygg");
            progtriangle.Add("Byggingenjör", "Energi och bygg");
            progtriangle.Add("Byggteknik - Produktion och förvaltning", "Energi och bygg");
            progtriangle.Add("Digitala tjänster 120 hp", "Data och IT");
            progtriangle.Add("Ekonomprogrammet", "Ekonomi och arbetsliv");
            progtriangle.Add("Ekonomprogrammet - Business Management", "Ekonomi och arbetsliv");
            progtriangle.Add("Entreprenöriellt företagande", "Ekonomi och arbetsliv");
            progtriangle.Add("Film- och TV produktion", "Medier");
            progtriangle.Add("Global sexuell och reproduktiv hälsa", "Vård");
            progtriangle.Add("Grafisk design", "Data och IT");
            progtriangle.Add("Grundlärare årskurs 4-6 arbetsintegrerad", "Lärarutb");
            progtriangle.Add("Grundlärare Förskoleklass - årskurs 1-3", "Lärarutb");
            progtriangle.Add("Grundlärare Grundskolans årskurs 4-6", "Lärarutb");
            progtriangle.Add("Högskoleingenjör hållbara energisystem", "Energi och bygg");
            progtriangle.Add("Industriell ekonomi - Högskoleexamen", "Industri");
            progtriangle.Add("International Tourism Management", "Ekonomi och arbetsliv");
            progtriangle.Add("IT säkerhet och mjukvarutestning", "Data och IT");
            progtriangle.Add("Kompl lärarutb år 7-9", "Lärarutb");
            progtriangle.Add("Kompl lärarutb gymnasiet", "Lärarutb");
            progtriangle.Add("Ljud- och musikproduktionsprogrammet", "Medier");
            progtriangle.Add("Magisterpgm Business Intelligence", "Ekonomi och arbetsliv");
            progtriangle.Add("Magisterpgm Demokrati, medborgarskap och förändring", "Samhälle");
            progtriangle.Add("Magisterpgm i Audiovisuella Studier", "Medier");
            progtriangle.Add("Magisterpgm i destinationsutveckling", "Ekonomi och arbetsliv");
            progtriangle.Add("Magisterpgm i Energieffektivt byggande", "Energi och bygg");
            progtriangle.Add("Magisterpgm i engelska - eng litteratur", "Språk");
            progtriangle.Add("Magisterpgm i engelska - lingvistik", "Språk");
            progtriangle.Add("Magisterpgm i Företagsekonomi", "Ekonomi och arbetsliv");
            progtriangle.Add("Magisterpgm i fysioterapi", "Vård");
            progtriangle.Add("Magisterpgm i Nationalekonomi", "Ekonomi och arbetsliv");
            progtriangle.Add("Magisterpgm i Pedagogiskt arbete", "Lärarutb");
            progtriangle.Add("Magisterpgm i solenergiteknik", "Energi och bygg");
            progtriangle.Add("Magisterpgm Svenska som andraspråk", "Språk");
            progtriangle.Add("Manus för film och TV", "Medier");
            progtriangle.Add("Masterprogram i Data Science", "Data och IT");
            progtriangle.Add("Masterprogram i interkulturella litteraturstudier", "Språk");
            progtriangle.Add("Masterprogram i solenergiteknik", "Energi och bygg");
            progtriangle.Add("Musik- och ljuddesign", "Medier");
            progtriangle.Add("Personal- och arbetslivsprogram", "Ekonomi och arbetsliv");
            progtriangle.Add("Produktionstekniker 120 hp", "Industri");
            progtriangle.Add("Samhällsvetarprogrammet", "Samhälle");
            progtriangle.Add("Sjuksköterskeprogrammet", "Vård");
            progtriangle.Add("Socionomprogrammet", "Socionom");
            progtriangle.Add("Spec.sjuksköterska - demensvård", "Vård");
            progtriangle.Add("Spec.sjuksköterska - distriktsköterska", "Vård");
            progtriangle.Add("Spec.sjuksköterska - vård av äldre", "Vård");
            progtriangle.Add("Sport Management", "Ekonomi och arbetsliv");
            progtriangle.Add("Systemvetenskapliga programmet", "Data och IT");
            progtriangle.Add("Tekniskt basår", "IIT resurs");
            progtriangle.Add("Tränarprogrammet för hälsa och idrott", "Lärarutb");
            progtriangle.Add("Vidareutbildning av lärare", "Lärarutb");
            progtriangle.Add("Yrkeslärare", "Lärarutb");
            progtriangle.Add("Byggarbetsledarprogrammet", "Energi och bygg");
            progtriangle.Add("Detaljhandelprogrammet 120hp", "Ekonomi och arbetsliv");
            progtriangle.Add("Detaljhandelsprogrammet 180hp", "Ekonomi och arbetsliv");
            progtriangle.Add("Digitalbrott och eSäkerhet", "Data och IT");
            progtriangle.Add("Energiteknikerprogrammet", "Energi och bygg");
            progtriangle.Add("Filmdesign för reklam och information", "Medier");
            progtriangle.Add("Filmproduktion 180 hp", "Medier");
            progtriangle.Add("Förskollärarprogrammet", "Lärarutb");
            progtriangle.Add("Grundlärare årskurs F-3 arbetsintegrerad", "Lärarutb");
            progtriangle.Add("Högskoleingenjör - Maskinteknik", "Industri");
            progtriangle.Add("Idrottstränarpgm m. hälsoinriktning", "Idrott och hälsa");
            progtriangle.Add("Idrottstränarpgm. m. inr mot prestation", "Idrott och hälsa");
            progtriangle.Add("Idrottstränarprogrammet - prestation och hälsa", "Idrott och hälsa");
            progtriangle.Add("Industriell ekonomi", "Industri");
            progtriangle.Add("Informationsdesign - kandidatprogram", "Data och IT");
            progtriangle.Add("Kompl lärarutb yrkes år 7-9", "Lärarutb");
            progtriangle.Add("Lärarprogrammet", "Lärarutb");
            progtriangle.Add("Magisterpgm  i religionsvetenskap", "IKS resurs");
            progtriangle.Add("Magisterpgm Maskin/Materialteknik", "Industri");
            progtriangle.Add("Manus och producentskap 180 hp", "Medier");
            progtriangle.Add("Master Business Intelligence", "Ekonomi och arbetsliv");
            progtriangle.Add("Masterprogram i Materialteknik och Produktutveckling", "Industri");
            progtriangle.Add("Medieproduktion för TV och sociala medier 180 hp", "Medier");
            progtriangle.Add("Musik- och ljudproduktion 180 hp", "Medier");
            progtriangle.Add("Teknisk/Naturvetenskaplig bastermin", "IIT resurs");
            progtriangle.Add("Upplevelseproduktion - friluftsliv, idrott och hälsa", "Idrott och hälsa");

            //foreach (string name in progtriangle.Keys.ToList())
            //{
            //    programclass pc = findprogrambyname(name);
            //    if (pc == null)
            //    {
            //        memo(name + " not found in progtriangle");
            //    }
            //    else if (!progtriangle.ContainsKey(pc.name))
            //    {
            //        progtriangle.Add(pc.name, progtriangle[name]);
            //    }

            //}

            subjcodetriangle.Add("VÅ", "Vård");
            subjcodetriangle.Add("VV", "Vård");
            subjcodetriangle.Add("MC", "Vård");
            subjcodetriangle.Add("IH", "Idrott och hälsa");
            subjcodetriangle.Add("SR", "Vård");
            subjcodetriangle.Add("SA", "Socionom");
            subjcodetriangle.Add("MI", "Data och IT");
            subjcodetriangle.Add("IK", "Data och IT");
            subjcodetriangle.Add("ST", "Data och IT");
            subjcodetriangle.Add("EG", "Energi och bygg");
            subjcodetriangle.Add("MT", "Industri");
            subjcodetriangle.Add("BY", "Energi och bygg");
            subjcodetriangle.Add("DT", "Data och IT");
            subjcodetriangle.Add("ET", "Industri");
            subjcodetriangle.Add("EU", "Ekonomi och arbetsliv");
            subjcodetriangle.Add("IE", "Industri");
            subjcodetriangle.Add("SK", "Samhälle");
            subjcodetriangle.Add("AB", "Ekonomi och arbetsliv");
            subjcodetriangle.Add("SO", "Samhälle");
            subjcodetriangle.Add("TR", "Ekonomi och arbetsliv");
            subjcodetriangle.Add("FÖ", "Ekonomi och arbetsliv");
            subjcodetriangle.Add("RK", "IKS resurs");
            subjcodetriangle.Add("HI", "IKS resurs");
            subjcodetriangle.Add("RV", "Ekonomi och arbetsliv");
            subjcodetriangle.Add("NA", "Ekonomi och arbetsliv");
            subjcodetriangle.Add("LP", "Medier");
            subjcodetriangle.Add("FI", "IKS resurs");
            subjcodetriangle.Add("BP", "Medier");
            subjcodetriangle.Add("PE", "Lärarutb");
            subjcodetriangle.Add("PG", "Lärarutb");
            subjcodetriangle.Add("NV", "Lärarutb");
            subjcodetriangle.Add("MD", "Lärarutb");
            subjcodetriangle.Add("AR", "Språk");
            subjcodetriangle.Add("JP", "Språk");
            subjcodetriangle.Add("SS", "Språk");
            subjcodetriangle.Add("SP", "Språk");
            subjcodetriangle.Add("IT", "Språk");
            subjcodetriangle.Add("FR", "Språk");
            subjcodetriangle.Add("LI", "Språk");
            subjcodetriangle.Add("SV", "Språk");
            subjcodetriangle.Add("TY", "Språk");
            subjcodetriangle.Add("EN", "Språk");
            subjcodetriangle.Add("PR", "Språk");
            subjcodetriangle.Add("RY", "Språk");
            subjcodetriangle.Add("KI", "Språk");
            subjcodetriangle.Add("FY", "IIT resurs");
            subjcodetriangle.Add("MA", "IIT resurs");
            subjcodetriangle.Add("KE", "IIT resurs");
            subjcodetriangle.Add("BQ", "Medier");
            subjcodetriangle.Add("PA", "Ekonomi och arbetsliv");
            subjcodetriangle.Add("PS", "Samhälle");
            subjcodetriangle.Add("AS", "IKS resurs");
            subjcodetriangle.Add("AU", "Medier");
            subjcodetriangle.Add("BI", "Lärarutb");
            subjcodetriangle.Add("MÖ", "Energi och bygg");
            subjcodetriangle.Add("GT", "Data och IT");
            subjcodetriangle.Add("KG", "Ekonomi och arbetsliv");
            subjcodetriangle.Add("KT", "IIT resurs");
            subjcodetriangle.Add("MP", "Industri");
            subjcodetriangle.Add("SQ", "Energi och bygg");



        }
        private void resapplbutton_Click(object sender, EventArgs e)
        {
            Dictionary<string, double> fkacc = new Dictionary<string, double>();
            Dictionary<string, double> fkres = new Dictionary<string, double>();
            Dictionary<string, double> progacc = new Dictionary<string, double>();
            Dictionary<string, double> progres = new Dictionary<string, double>();

            fill_triangledicts();

            string fn = @"C:\Users\lsj\OneDrive - Högskolan Dalarna\Dokument\Invärld\Utbildningskonsolidering 2024\Reserver per antagen Johan 1flik.txt";
            using (StreamReader sr = new StreamReader(fn))
            {
                sr.ReadLine();
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    string[] words = line.Split('\t');
                    bool fkprog = (words[1] == "fk");
                    string triangle;
                    if (fkprog)
                        triangle = subjectriangle[words[2]];
                    else
                        triangle = progtriangle[words[2]];
                    if (!fkacc.ContainsKey(triangle))
                    {
                        fkacc.Add(triangle, 0);
                        fkres.Add(triangle, 0);
                        progacc.Add(triangle, 0);
                        progres.Add(triangle, 0);
                    }

                    double acc = util.tryconvert(words[3]);
                    double res = util.tryconvert(words[4]);

                    if (fkprog)
                    {
                        fkacc[triangle] += acc;
                        fkres[triangle] += res;
                    }
                    else
                    {
                        progacc[triangle] += acc;
                        progres[triangle] += res;
                    }
                }
            }

            memo("triangel,fkresacc,progresacc");
            foreach (string tri in fkacc.Keys)
            {
                double fkresacc = fkacc[tri] > 0 ? fkres[tri] / fkacc[tri] : 0;
                double progresacc = progacc[tri] > 0 ? progres[tri] / progacc[tri] : 0;
                memo(tri + "," + fkresacc.ToString("N2", CultureInfo.InvariantCulture) + "," + progresacc.ToString("N2", CultureInfo.InvariantCulture));
            }

            memo("triangel\tfkresacc\tprogresacc\tfkacc\tprogacc");
            foreach (string tri in fkacc.Keys)
            {
                double fkresacc = fkacc[tri] > 0 ? fkres[tri] / fkacc[tri] : 0;
                double progresacc = progacc[tri] > 0 ? progres[tri] / progacc[tri] : 0;
                memo(tri + "\t" + fkresacc.ToString("N2") + "\t" + progresacc.ToString("N2")+"\t"+fkacc[tri]+"\t"+progacc[tri]);
            }
        }
    }
}
