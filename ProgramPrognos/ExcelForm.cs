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
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProgramPrognos
{
    public partial class ExcelForm : Form
    {
        string lastsemwithdata;
        public ExcelForm()
        {
            InitializeComponent();
        }
        public void memo(string s)
        {
            richTextBox1.AppendText(s + "\n");
            richTextBox1.ScrollToCaret();
        }

        private void SheetWithHeader(Excel.Worksheet sheet, int datarows, Dictionary<string,int> hd)
        {
            for (int i = 0; i <= datarows; i++)
                sheet.Rows.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);
            foreach (string hh in hd.Keys)
            {
                sheet.Cells[1, hd[hh] + 1] = hh;
            }
            //Excel.Range qa = sheet.Columns[hd[pubclass.authorstring] + 1];
            //qa.ColumnWidth = 50;
            //Excel.Range qt = sheet.Columns[hd[pubclass.titstring] + 1];
            //qt.ColumnWidth = 50;
            //sheet.Columns[pubclass.titstring + 1].ColumnWidth = 300;
            //sheet.Cells[1, pubclass.titstring + 1].EntireColumn.ColumnWidth = 300;
            //Excel.Range titcol = ((Excel.Range)sheet.Cells[1, pubclass.titstring+1]).EntireColumn;
            //titcol.ColumnWidth = 300;
            //Excel.Range aucol = ((Excel.Range)sheet.Cells[1, pubclass.authorstring + 1]).EntireColumn;
            //titcol.ColumnWidth = 400;
        }

        static int beforeA = (int)'A' - 1;

        private string Cellname(int row,int col)
        {
            
            int nalph = (col-1) / 26;
            int nlett = col % 26;
            if (nlett == 0)
                nlett = 26;
            string s;
            if (nalph == 0)
                s = "";
            else
                s = ((char)(beforeA + nalph)).ToString();
            s += (char)(beforeA + nlett)  + row.ToString();
            return s;
        }
        private Dictionary<string, int> ProgramNames(Excel.Worksheet sheet, List<programclass> qprog) //list program names in column A of sheet
        {
            return ProgramNames(sheet, qprog, 1);
        }

        private Dictionary<string,int> ProgramNames(Excel.Worksheet sheet, List<programclass> qprog,int nheadrows) //list program names in column A of sheet
        {
            Dictionary<string, int> dict = new Dictionary<string, int>();
            int nrow = nheadrows;
            foreach (programclass pc in qprog)
            {
                nrow++;
                sheet.Cells[nrow, 1] = pc.name;
                dict.Add(pc.name, nrow);
                //sheet.Cells[nrow, 2].Formula = "=UPPER(A" + nrow + ")";

                //Excel.Range rr;

            }
            Excel.Range qa = sheet.Columns[1];
            qa.ColumnWidth = 50;
            return dict;
        }

        private void RetentionSheet(Excel.Worksheet retsheet, List<programclass> qprog, int allmaxsem)
        {
            rethd = new Dictionary<string, int>() { { "Program", 0 }, { "Medelretention", 1 }, { "Från antagen till registrerad", 2 } };
            for (int i = 1; i < allmaxsem; i++)
            {
                rethd.Add("T" + i + "->T" + (i + 1), 2 + i);
            }

            SheetWithHeader(retsheet, qprog.Count + 1, rethd);
            prow = ProgramNames(retsheet, qprog,2);

            int offset = 3;
            int meancol = 2;

            foreach (programclass pc in qprog)
            {
                int row = prow[pc.name];
                for (int i=0;i<pc.semesters;i++)
                {
                    retsheet.Cells[row, i + offset] = pc.transition[i].transitionprob;
                }
                retsheet.Cells[row, meancol].Formula = "=AVERAGE(" + Cellname(row, offset) + ":" + Cellname(row, offset + pc.semesters) + ")";
            }
            retsheet.Range["B2", Cellname(qprog.Count + 2, allmaxsem + 3)].NumberFormat = "###.0%";
            retsheet.Cells[1, 1].Locked = false;
            retsheet.Protect();
        }

        Dictionary<string, int> prow;
        Dictionary<string, int> crow;
        Dictionary<string, int> planhd;
        Dictionary<string, int> plan2hd;
        Dictionary<string, int> bathd;
        Dictionary<string, int> rethd;
        Dictionary<string, int> coursehd;
        int retoffset = 2;
        int courseoffset = 2;
        string acceptstring = "Antas ";
        string t1string = "T1 ";
        string studstring = "Stud ";
        string fkstudstring = "FK-stud ";
        string fkstring = "FK ";
        string progstudstring = "Prog-stud ";
        string progstring = "prog ";
        string hststring = "HST ";
        string hprstring = "HPR ";
        string moneystring = "Kr ";
        string retsheetname = "Retention";
        string mainsheetname = "Planering";
        string detailsheetname = "Detaljer";
        string batsheetname = "Programkullar";
        string coursesheetname = "Kurser";
        string paysheetname = "Betalande stud";
        string applstring = "1:ahandssökande";
        string accstring = "Antagna";
        string studhststring = "Stud/HST?";
        string prestationstring = "Prest-grad";

        private void fill_planhd(Excel.Worksheet sheet, Excel.Worksheet sheet2, List<programclass> qprog, string startsem, string endsem)
        {
            planhd = new Dictionary<string, int>() { { "Program", 0 } };
            plan2hd = new Dictionary<string, int>() { { "Program", 0 }, { "Medelretention", 1 }, { "Från antagen till registrerad", 2 }, { "Prestationsgrad", 3 } };
            List<string> semlist = new List<string>();

            int startyear = 2000 + util.semtoint(startsem);
            int endyear = 2000 + util.semtoint(endsem);

            int col = planhd.Count;
            int col2 = plan2hd.Count;
            for (int i = startyear; i <= endyear; i++)
            {
                plan2hd.Add(hststring + i, col2);
                col2++;
                plan2hd.Add(hprstring + i, col2);
                col2++;
                planhd.Add(moneystring + i, col);
                col++;
            }

            planhd.Add(studhststring, col);
            col++;

            string sem = startsem;
            //col = plan2hd.Count + 1;
            do
            {
                plan2hd.Add(acceptstring + sem, col2);
                semlist.Add(sem);
                col2++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));

            sem = startsem;
            //col = planhd.Count + 1;
            do
            {
                planhd.Add(t1string + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));

            sem = startsem;
            //col = planhd.Count + 1;
            do
            {
                plan2hd.Add(hststring + sem, col2);
                col2++;
                plan2hd.Add(hprstring + sem, col2);
                col2++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));

            SheetWithHeader(sheet, qprog.Count + 2, planhd);
            SheetWithHeader(sheet2, qprog.Count + 2, plan2hd);

            //sheet.Protect();

        }

        private void PlanSheet(Excel.Worksheet sheet, List<programclass> qprog, string startsem, string endsem,string inst)
        {
            Dictionary<string, int> prow = ProgramNames(sheet, qprog,2);
            Dictionary<string, bool> phtstart = new Dictionary<string, bool>();
            Dictionary<string, bool> pvtstart = new Dictionary<string, bool>();
            sheet.Cells[2, 1] = "Total";
            sheet.Rows[2].Font.Bold = true;
            for (int icol = 2; icol <= planhd.Count; icol++)
            {
                sheet.Cells[2, icol].Formula = "=SUM(" + Cellname(3, icol) + ":" + Cellname(3 + qprog.Count, icol) + ")";
            }




            //='Sheet 1'!A3

            int startyear = util.semtoint(startsem);
            int endyear = util.semtoint(endsem);

            double roundfactor = 5;

            int lastcolwithdata = -1;

            foreach (programclass pc in qprog)
            {
                int row = prow[pc.name];
                for (int year = startyear; year <= endyear; year++)
                {
                    int colyear = planhd[moneystring + "20" + year] + 1;
                    double hstpeng = pc.fracproddict[inst].hstpeng;
                    if (hstpeng == 0)
                        hstpeng = qprog.First().fracproddict[inst].hstpeng;
                    double hprpeng = pc.fracproddict[inst].hprpeng;
                    if (hprpeng == 0)
                        hprpeng = qprog.First().fracproddict[inst].hprpeng;
                    if (pc.fk)
                    {
                        int colvt = planhd[t1string + "VT" + year] + 1;
                        int colht = planhd[t1string + "HT" + year] + 1;
                        sheet.Cells[row, colyear].Formula = toreplace + "=" + (hstpeng+pc.totalprod.prestationsgrad()*hprpeng) + "*(" + Cellname(row, colvt) + "+" + Cellname(row,colht)+")";
                    }
                    else
                    {
                        int colhst = plan2hd[hststring + "20" + year] + 1;
                        int colhpr = plan2hd[hprstring + "20" + year] + 1;
                        sheet.Cells[row, colyear].Formula = toreplace + "=" + hstpeng + "*'" + detailsheetname + "'!" + Cellname(row, colhst) + "+" + hprpeng + "*'" + detailsheetname + "'!" + Cellname(row, colhpr);
                    }
                }

                string studhst = pc.fk ? "HST" : "Stud";
                sheet.Cells[row, planhd[studhststring] + 1] = studhst;

                phtstart.Add(pc.name, false);
                pvtstart.Add(pc.name, false);

                string sem = startsem;
                do
                {
                    int col = planhd[t1string + sem] + 1;
                    lastcolwithdata = col;
                    if (pc.fk)
                    {
                        double hst = pc.totalprod.frachst;
                        if (hst > 0)
                        {
                            sheet.Cells[row, col] = roundfactor * Math.Round(0.5 * hst / roundfactor);
                            phtstart[pc.name] = true;
                            pvtstart[pc.name] = true;
                        }
                    }
                    else
                    {
                        programbatchclass bc = (from c in pc.batchlist where c.batchstart == sem select c).FirstOrDefault();
                        if (bc != null)
                        {
                            sheet.Cells[row, col] = bc.getstud(1);
                            if (bc.batchstart.Contains("H"))
                                phtstart[pc.name] = true;
                            else
                                pvtstart[pc.name] = true;
                        }
                    }
                    sem = util.incrementsemester(sem);
                }
                while (sem != util.incrementsemester(lastsemwithdata));

            }
            //sheet.Range["B2", Cellname(qprog.Count + 1, allmaxsem + 2)].NumberFormat = "###.0%";
            //sheet.Cells[1, 1].Locked = false;
            sheet.Range["B2", Cellname(qprog.Count + 2, 6)].NumberFormat = "### ### ###";
            sheet.Range["B2", Cellname(qprog.Count + 2, 6)].Interior.Color = Excel.XlRgbColor.rgbLightPink;
            for (int i = 2; i < 7; i++)
            {
                Excel.Range qa = sheet.Columns[i];
                qa.ColumnWidth = 15;
                //qa.Interior.Color = Excel.XlRgbColor.rgbPink;
                //qa.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlLess, 10);
            }


            //for (int i = lastcolwithdata+1; i <=planhd.Count; i++)
            //{
            //    Excel.Range qa = sheet.Columns[i];
            //    //qa.ColumnWidth = 15;
            //    qa.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            //}

            sheet.Range[Cellname(3, lastcolwithdata + 1), Cellname(qprog.Count + 2, planhd.Count + 1)].Locked = false;
            sheet.Range[Cellname(3, lastcolwithdata + 1), Cellname(qprog.Count + 2, planhd.Count)].Interior.Color = Excel.XlRgbColor.rgbPaleGoldenrod;
            sheet.Range[Cellname(3, lastcolwithdata + 1), Cellname(qprog.Count + 2, planhd.Count)].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            foreach (programclass pc in qprog)
            {
                int row = prow[pc.name];
                bool ht = phtstart[pc.name];
                bool vt = pvtstart[pc.name];
                int icol = lastcolwithdata + 1;
                while (icol <= planhd.Count )
                {
                    if (ht)
                        sheet.Cells[row, icol].Interior.Color = Excel.XlRgbColor.rgbYellow;
                    if (vt && icol < planhd.Count)
                        sheet.Cells[row, icol+1].Interior.Color = Excel.XlRgbColor.rgbYellow;
                    icol += 2;
                }

            }

            //sheet.FreezeColumns(1);
            //sheet.Protect();
        }

        string toreplace = "§§§";

        private void DetailSheet(Excel.Worksheet sheet, List<programclass> qprog, string startsem, string endsem,string inst)
        {
            Dictionary<string, int> prow = ProgramNames(sheet, qprog, 2);
            sheet.Cells[2, 1] = "Total";
            sheet.Rows[2].Font.Bold = true;
            for (int icol = 5;icol<=plan2hd.Count;icol++)
            {
                sheet.Cells[2,icol].Formula = toreplace+"=SUM("+Cellname(3,icol)+":"+Cellname(3+qprog.Count,icol)+")";
            }

            int meancol = 2;
            int tr0col = 3;
            int prestcol = 4;

            //='Sheet 1'!A3

            foreach (programclass pc in qprog)
            {
                int row = prow[pc.name];
                sheet.Cells[row, meancol].Formula = "='" + retsheetname + "'!" + Cellname(row, meancol);
                sheet.Cells[row, tr0col].Formula = "='" + retsheetname + "'!" + Cellname(row, tr0col);
                double prest = pc.prod_per_student.prestationsgrad();
                if (prest > 1)
                    prest = 0.8;
                sheet.Cells[row, prestcol] = prest;

                int batrow = 300;
                //double frachst = 0.5; // pc.fracproddict[inst].frachst;
                double frachst = pc.fracproddict[inst].frachst;
                //double frachpr = frachst*0.8; // pc.fracproddict[inst].frachst;
                double frachpr = pc.fracproddict[inst].frachpr;
                string semx = startsem;
                do
                {
                    //int retcol = retoffset + nsem;
                    string hsts = "=SUMIF('" + batsheetname + "'!A2:Z" + batrow + ";" + Cellname(row, 1) + ";'" + batsheetname + "'!" + Cellname(2, bathd[semx] + 1) + ":" + Cellname(batrow, bathd[semx] + 1) + ")*" + frachst; //"*'" + retsheetname + "'!" + Cellname(prow[prog], retcol);
                    sheet.Cells[row, plan2hd[hststring + semx] + 1] = toreplace+hsts;
                    //sheet.Cells[row, plan2hd[hststring + semx] + 1] = hsts;
                    string hprs = "=SUMIF('" + batsheetname + "'!A2:Z" + batrow + ";" + Cellname(row, 1) + ";'" + batsheetname + "'!" + Cellname(2, bathd[semx] + 1) + ":" + Cellname(batrow, bathd[semx] + 1) + ")*" + frachpr; //"*'" + retsheetname + "'!" + Cellname(prow[prog], retcol);
                    sheet.Cells[row, plan2hd[hprstring + semx] + 1] = toreplace+hprs;

                    string antags = "='" + mainsheetname + "'!" + Cellname(row, planhd[t1string + semx] + 1)+"/'"+retsheetname+"'!"+Cellname(row,3);
                    sheet.Cells[row, plan2hd[acceptstring + semx]+1].Formula = antags;
                    semx = util.incrementsemester(semx);

                }
                while (semx != util.incrementsemester(endsem));

                int startyear = util.semtoint(startsem);
                int endyear = util.semtoint(endsem);
                for (int year = startyear; year <= endyear; year++)
                {
                    int colyear = plan2hd[hststring + "20" + year]+1;
                    int colvt = plan2hd[hststring + "VT" + year] + 1;
                    int colht = plan2hd[hststring + "HT" + year] + 1;
                    sheet.Cells[row, colyear].Formula = toreplace+"=" + Cellname(row, colvt) + "+" + Cellname(row, colht);
                    colyear = plan2hd[hprstring + "20" + year] + 1;
                    colvt = plan2hd[hprstring + "VT" + year] + 1;
                    colht = plan2hd[hprstring + "HT" + year] + 1;
                    sheet.Cells[row, colyear].Formula = toreplace+"=" + Cellname(row, colvt) + "+" + Cellname(row, colht);
                }
            }

            //= SUMIF(Programkullar!A1: W41; A2; Programkullar!H1: H99)


            //sheet.Range["B2", Cellname(qprog.Count + 1, allmaxsem + 2)].NumberFormat = "###.0%";
            sheet.Cells[1, 1].Locked = false;
            sheet.Range["B2", Cellname(qprog.Count + 2, 4)].NumberFormat = "###.0%";
            sheet.Range["e2", Cellname(qprog.Count + 2, plan2hd.Count+1)].NumberFormat = "# ###.#";

            foreach (string s in plan2hd.Keys)
            {
                if (s.Contains(acceptstring))
                {
                    Excel.Range qa = sheet.Columns[plan2hd[s] + 1];
                    qa.ColumnWidth = 11;
                    qa.NumberFormat = "# ###";
                }

            }

            //sheet.FreezeColumns(1);
            //sheet.Protect();
        }

        private void CourseSheet(Excel.Worksheet coursesheet, List<programclass> qprog, List<programclass> qcourse, string startsem, string endsem)
        {
            int hpcol = 2;
            int codecol = 3;
            int prestcol = 4;
            int moneycol = 5;
            int totalrow = 2;

            int newcourses = 20;
            string newcoursename = "Ny kurs ";

            coursehd = new Dictionary<string, int>() { 
                { "Kurs", 0 }, 
                { "Hp", hpcol-1 },
                { "Kurskod", codecol-1 }, 
                { prestationstring, prestcol-1 }, 
                { "Kr/HST", moneycol-1 } };

            List<string> semlist = new List<string>();
            int col = coursehd.Count + 1;

            int startyear = 2000 + util.semtoint(startsem);
            int endyear = 2000 + util.semtoint(endsem);

            for (int i = startyear; i <= endyear; i++)
            {
                coursehd.Add(moneystring + fkstring + i, col);
                col++;
            }

            for (int i = startyear; i <= endyear; i++)
            {
                coursehd.Add(moneystring + progstring + i, col);
                col++;
            }


            string sem = startsem;

            do
            {
                coursehd.Add(studstring + sem, col);
                semlist.Add(sem);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));

            sem = startsem;

            do
            {
                coursehd.Add(fkstudstring + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));

            sem = startsem;

            do
            {
                coursehd.Add(progstudstring + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));

            sem = startsem;

            do
            {
                coursehd.Add(acceptstring + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));


            sem = startsem;

            do
            {
                coursehd.Add(hststring + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));


            sem = startsem;

            do
            {
                coursehd.Add(hprstring + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));


            sem = startsem;

            do
            {
                coursehd.Add(hststring + fkstring + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));


            sem = startsem;

            do
            {
                coursehd.Add(hprstring + fkstring + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));

            sem = startsem;

            do
            {
                coursehd.Add(hststring + progstring + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));

            sem = startsem;

            do
            {
                coursehd.Add(hprstring + progstring + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));

            for (int i = startyear; i <= endyear; i++)
            {
                coursehd.Add(hststring + i, col);
                col++;
                coursehd.Add(hprstring + i, col);
                col++;
            }

            //for (int i = 1; i < allmaxsem; i++)
            //{
            //    rethd.Add("T" + i + "->T" + (i + 1), 2 + i);
            //}

            Dictionary<string, List<string>> progrefdict = new Dictionary<string, List<string>>();
            foreach (programclass pc in qprog)
            {
                foreach (int i in pc.coursedict.Keys)
                {
                    foreach (string cc in pc.coursedict[i].Keys)
                    {
                        if (!progrefdict.ContainsKey(cc))
                        {
                            progrefdict.Add(cc, new List<string>());
                        }
                        if (!progrefdict[cc].Contains(pc.name))
                            progrefdict[cc].Add(pc.name);
                    }
                }
            }

            for (int i=0;i<newcourses;i++)
            {
                programclass nc = new programclass(newcoursename + i + "(fyll i hp och FK-stud)");
                qcourse.Add(nc);
            }

            SheetWithHeader(coursesheet, qcourse.Count + courseoffset, coursehd);
            crow = ProgramNames(coursesheet, qcourse, courseoffset);
            coursesheet.Cells[totalrow, 1] = "Total";
            coursesheet.Rows[totalrow].Font.Bold = true;

            double prestsum = 0;
            double krhstsum = 0;
            double ncourses = 0;


            foreach (programclass pc in qcourse)
            {
                bool nykurs = pc.name.StartsWith(newcoursename);
                int nrow = crow[pc.name];
                if (!nykurs)
                    coursesheet.Cells[nrow, hpcol].Value = pc.hp;
                coursesheet.Cells[nrow, codecol].Value = pc.bestcode();
                double prest = 0.8;
                if (pc.totalprod.frachst > 0)
                {
                    prest = pc.totalprod.frachpr / pc.totalprod.frachst;
                    if (prest > 1)
                        prest = 1;
                }
                else if (nykurs)
                {
                    prest = prestsum / ncourses;
                }
                prestsum += prest;
                coursesheet.Cells[nrow, prestcol].Value = prest;
                double krhst;
                if (nykurs)
                    krhst = krhstsum / ncourses;
                else
                    krhst = Form1.hstkr(1, pc.studentpengarea) + prest * Form1.hprkr(1, pc.studentpengarea);
                coursesheet.Cells[nrow, moneycol].Value = krhst;
                krhstsum += krhst;
                ncourses++;

                //FK-studenter:
                double lastvt = 0;
                double lastht = 0;
                foreach (string sm in semlist)
                {
                    int ncol = coursehd[fkstudstring + sm] + 1;
                    programbatchclass bc = pc.getbatch(sm);
                    if (bc != null)
                    {
                        double stud = (double)bc.getactualstud(1);
                        coursesheet.Cells[nrow, ncol].Value = stud;
                        if (sm.StartsWith("V"))
                            lastvt = stud;
                        else
                            lastht = stud;
                    }
                    else
                    {
                        double stud;
                        if (sm.StartsWith("V"))
                            stud = lastvt;
                        else
                            stud = lastht;
                        if (stud > 0)
                            coursesheet.Cells[nrow, ncol].Value = stud;
                    }

                }

                //programstudenter:
                if (progrefdict.ContainsKey(pc.bestcode()))
                {
                    foreach (string sm in semlist)
                    {
                        int ncol = coursehd[progstudstring + sm] + 1;
                        string sumref = "";
                        foreach (string prog in progrefdict[pc.bestcode()])
                        {
                            programclass prc = Form1.origprogramdict[prog];
                            foreach (int isem in prc.coursedict.Keys)
                            {
                                if (prc.coursedict[isem].ContainsKey(pc.bestcode()))
                                {
                                    if (prc.coursedict[isem][pc.bestcode()] > 0)
                                    {
                                        string batstart = util.find_batstart(sm, isem);
                                        if (batsemref[prog].ContainsKey(batstart)
                                            && batsemref[prog][batstart].ContainsKey(isem))
                                        {
                                            if (String.IsNullOrEmpty(sumref))
                                                sumref = toreplace + "=";
                                            else
                                                sumref += "+";
                                            sumref += batsemref[prog][batstart][isem];
                                            sumref += "*" + prc.coursedict[isem][pc.bestcode()];
                                        }
                                    }
                                }
                            }
                        }
                        coursesheet.Cells[nrow, ncol].Value = sumref;
                    }
                }

                //summa studenter
                foreach (string sm in semlist)
                {
                    int ncol = coursehd[studstring + sm] + 1;
                    int ncolf = coursehd[fkstudstring + sm] + 1;
                    int ncolp = coursehd[progstudstring + sm] + 1;
                    coursesheet.Cells[nrow, ncol].Formula = toreplace + "=IF("
                        + Cellname(nrow, ncolf) + "+" + Cellname(nrow, ncolp) + "=0;\"\";"
                        + Cellname(nrow, ncolf) + "+" + Cellname(nrow, ncolp) + ")";
                }

                //HST FK
                foreach (string sm in semlist)
                {
                    int ncol = coursehd[hststring+fkstring + sm] + 1;
                    int ncolf = coursehd[fkstudstring + sm] + 1;
                    
                    coursesheet.Cells[nrow, ncol].Formula = toreplace+"=" + Cellname(nrow, ncolf)
                        + "*" + Cellname(nrow, hpcol)+"/60";
                }

                //HST Prog
                foreach (string sm in semlist)
                {
                    int ncol = coursehd[hststring + progstring + sm] + 1;
                    int ncolf = coursehd[progstudstring + sm] + 1;

                    coursesheet.Cells[nrow, ncol].Formula = toreplace + "=" + Cellname(nrow, ncolf)
                        + "*" + Cellname(nrow, hpcol) + "/60";
                }

                //summa HST
                foreach (string sm in semlist)
                {
                    int ncol = coursehd[hststring + sm] + 1;
                    int ncolf = coursehd[hststring+fkstring + sm] + 1;
                    int ncolp = coursehd[hststring+progstring + sm] + 1;
                    coursesheet.Cells[nrow, ncol].Formula = toreplace+"=" + Cellname(nrow, ncolf) + "+" + Cellname(nrow, ncolp);
                }


                //HPR FK
                foreach (string sm in semlist)
                {
                    int ncol = coursehd[hprstring + fkstring + sm] + 1;
                    int ncolf = coursehd[hststring+fkstring + sm] + 1;

                    coursesheet.Cells[nrow, ncol].Formula = toreplace + "=" + Cellname(nrow, ncolf)
                        + "*" + Cellname(nrow, prestcol);
                }

                //HPR prog
                foreach (string sm in semlist)
                {
                    int ncol = coursehd[hprstring + progstring + sm] + 1;
                    int ncolf = coursehd[hststring + progstring + sm] + 1;

                    coursesheet.Cells[nrow, ncol].Formula = toreplace + "=" + Cellname(nrow, ncolf)
                        + "*" + Cellname(nrow, prestcol);
                }

                //summa HPR
                foreach (string sm in semlist)
                {
                    int ncol = coursehd[hprstring + sm] + 1;
                    int ncolf = coursehd[hprstring + fkstring + sm] + 1;
                    int ncolp = coursehd[hprstring + progstring + sm] + 1;
                    coursesheet.Cells[nrow, ncol].Formula = toreplace + "=" + Cellname(nrow, ncolf) + "+" + Cellname(nrow, ncolp);
                }

                //Pengar FK
                for (int i = startyear; i <= endyear; i++)
                {
                    int ncol = coursehd[moneystring + fkstring + i]+1;
                    string vtsem = "VT" + (i % 100);
                    string htsem = "HT" + (i % 100);
                    int hstvtcol = coursehd[hststring + fkstring + vtsem] + 1;
                    int hsthtcol = coursehd[hststring + fkstring + htsem] + 1;
                    int hprvtcol = coursehd[hprstring + fkstring + vtsem] + 1;
                    int hprhtcol = coursehd[hprstring + fkstring + htsem] + 1;
                    string f = toreplace + "=";
                    if (nykurs)
                    {
                        f += Cellname(nrow, moneycol) + "*(" + Cellname(nrow, hstvtcol) + "+" + Cellname(nrow, hsthtcol) + ")";
                    }
                    else
                    {
                        f += Form1.hstkr(1, pc.studentpengarea) + "*(" + Cellname(nrow, hstvtcol) + "+" + Cellname(nrow, hsthtcol) + ")+";
                        f += Form1.hprkr(1, pc.studentpengarea) + "*(" + Cellname(nrow, hprvtcol) + "+" + Cellname(nrow, hprhtcol) + ")";
                    }
                    coursesheet.Cells[nrow, ncol].Formula = f;

                    col++;
                }

                //Pengar prog
                for (int i = startyear; i <= endyear; i++)
                {
                    int ncol = coursehd[moneystring + progstring + i] + 1;
                    string vtsem = "VT" + (i % 100);
                    string htsem = "HT" + (i % 100);
                    int hstvtcol = coursehd[hststring + progstring + vtsem] + 1;
                    int hsthtcol = coursehd[hststring + progstring + htsem] + 1;
                    int hprvtcol = coursehd[hprstring + progstring + vtsem] + 1;
                    int hprhtcol = coursehd[hprstring + progstring + htsem] + 1;
                    string f = toreplace + "=";
                    f += Form1.hstkr(1, pc.studentpengarea) + "*(" + Cellname(nrow, hstvtcol) + "+" + Cellname(nrow, hsthtcol) + ")+";
                    f += Form1.hprkr(1, pc.studentpengarea) + "*(" + Cellname(nrow, hprvtcol) + "+" + Cellname(nrow, hprhtcol) + ")";
                    coursesheet.Cells[nrow, ncol].Formula = f;

                    col++;
                }

                //Antas
                foreach (string sm in semlist)
                {
                    int ncol = coursehd[acceptstring + sm] + 1;
                    int ncolf = coursehd[fkstudstring + sm] + 1;
                    double transition = 0.666;
                    if (pc.transition[0] != null && pc.transition[0].transitionprob > 0)
                        transition = pc.transition[0].transitionprob;
                    coursesheet.Cells[nrow, ncol].Formula = toreplace + "=IF(" + Cellname(nrow, ncolf) + ">0;" + Cellname(nrow, ncolf)+"/"+transition+";\"\")";
                }

            }
            //sdfa

            for (int icol=6;icol<=coursehd.Count;icol++)
            {
                coursesheet.Cells[totalrow,icol].Formula = toreplace+"=SUM("+Cellname(totalrow+1,icol)+":"+Cellname(qcourse.Count+courseoffset,icol)+")";
            }

            coursesheet.Range["D3", Cellname(999, 4)].NumberFormat = "###.0%";
            coursesheet.Range["g3", "p999"].NumberFormat = "# ### ###";
            coursesheet.Range["g3", "p999"].ColumnWidth = 13;
            //coursesheet.Range["f3", "O999"].Interior.Color = Color.Pink;
            ////coursesheet.Range["z3", "ai999"].NumberFormat = "# ###.#";
            //coursesheet.Range["z3", "ai999"].Interior.Color = Color.Yellow;
            coursesheet.Range["aj3", "dt999"].NumberFormat = "# ###.#";
            //coursesheet.Range["aj3", "as999"].Interior.Color = Color.Tan;
            foreach (string s in coursehd.Keys)
            {
                int ncol = coursehd[s] + 1;
                Color color = Color.White;
                if (s.Contains(moneystring + fkstring))
                    color = Color.Pink;
                else if (s.Contains(moneystring + progstring))
                    color = Color.LightPink;
                else if (s.Contains(fkstudstring))
                    color = Color.Yellow;
                else if (s.Contains(progstudstring))
                    color = Color.Tan;
                else if (s.Contains(acceptstring))
                    color = Color.LightGreen;
                else if (s.Contains(hststring))
                    color = Color.LightBlue;
                else if (s.Contains(hprstring))
                    color = Color.PaleTurquoise;
                else if (s.Contains(studstring))
                    color = Color.AntiqueWhite;
                coursesheet.Range[Cellname(3, ncol), Cellname(999, ncol)].Interior.Color = color;

                var qa = coursesheet.Range[Cellname(3, ncol), Cellname(999, ncol)];
                if (color == Color.Yellow)
                    qa.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                if(color == Color.AntiqueWhite)
                {
                    Excel.FormatCondition cond = qa.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlBetween,0.01, 9.99);
                    cond.Font.Color = Color.Red;
                    qa.NumberFormat = "###";
                }
            }

        }

        // batsemref[prog][batstart][sem] = Cellref.
        Dictionary<string, Dictionary<string, Dictionary<int, string>>> batsemref = new Dictionary<string, Dictionary<string, Dictionary<int, string>>>();

        private void add_batsemref(int nrow,int ncol,string prog, string bat, int sem,string shname)
        {
            string cname = "'" + shname + "'!"+Cellname(nrow, ncol);
            if (!batsemref.ContainsKey(prog))
                batsemref.Add(prog, new Dictionary<string, Dictionary<int, string>>());
            if (!batsemref[prog].ContainsKey(bat))
                batsemref[prog].Add(bat, new Dictionary<int, string>());
            if (!batsemref[prog][bat].ContainsKey(sem))
                batsemref[prog][bat].Add(sem, cname);
        }


        private void BatchSheet(Excel.Worksheet sheet, List<programclass> qprog, string startsem, string endsem)
        {
            bathd = new Dictionary<string, int>() { { "Program", 0 }, { "Start", 1 }, { applstring, 2 },{ accstring, 3 } };
            Dictionary<string, int> semdict = new Dictionary<string, int>();
            
            List<string> semlist = new List<string>();

            string sem = startsem;
            int col = 4;
            do
            {
                bathd.Add(sem, col);
                semlist.Add(sem);
                col++;
                sem = util.incrementsemester(sem);

            }
            while (sem != util.incrementsemester(endsem));

            int nrow = 0;
            Dictionary<string, Dictionary<string, Dictionary<string, double>>> progbatsem = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();
            Dictionary<string, Dictionary<string, double>> progbatappl = new Dictionary<string, Dictionary<string, double>>();
            Dictionary<string, Dictionary<string, double>> progbatacc = new Dictionary<string, Dictionary<string, double>>();
            foreach (programclass pc in qprog)
            {
                Dictionary<string, Dictionary<string, double>> batsem = new Dictionary<string, Dictionary<string, double>>();
                Dictionary<string, double> batappl = new Dictionary<string, double>();
                Dictionary<string, double> batacc = new Dictionary<string, double>();

                foreach (programbatchclass bc in pc.batchlist)
                {
                    Dictionary<string, double> dict = new Dictionary<string, double>();
                    foreach (string ss in semlist)
                    {
                        double stud = bc.getstud(ss);
                        if (stud > 0)
                            dict.Add(ss, stud);
                    }
                    if (dict.Count > 0)
                    {
                        batsem.Add(bc.batchstart, dict);
                        batappl.Add(bc.batchstart, bc.applicants);
                        batacc.Add(bc.batchstart, bc.getstud(0));
                        nrow++;
                    }
                }
                if (batsem.Count > 0)
                {
                    progbatsem.Add(pc.name, batsem);
                    progbatappl.Add(pc.name, batappl);
                    progbatacc.Add(pc.name, batacc);
                }

            }

            SheetWithHeader(sheet, nrow + 1, bathd);
            //Dictionary<string, int> prow = ProgramNames(sheet, qprog);

            nrow = 1;
            foreach (string prog in progbatsem.Keys)
            {
                int nsemtot = Form1.origprogramdict[prog].semesters;

                foreach (string bat in progbatsem[prog].Keys)
                {
                    nrow++;
                    sheet.Cells[nrow, bathd["Program"]+1] = prog;
                    sheet.Cells[nrow, bathd["Start"]+1] = bat;
                    sheet.Cells[nrow, bathd[applstring] + 1] = progbatappl[prog][bat];
                    sheet.Cells[nrow, bathd[accstring] + 1] = progbatacc[prog][bat];
                    foreach (string ss in progbatsem[prog][bat].Keys)
                    {
                        sheet.Cells[nrow, bathd[ss]+1] = progbatsem[prog][bat][ss];
                        // batsemref[prog][batstart][sem] = Cellref.
                        int isem = util.semestercount(bat, ss);
                        add_batsemref(nrow, bathd[ss] + 1, prog,bat, isem,batsheetname);

                    }
                    if (progbatsem[prog][bat].ContainsKey(lastsemwithdata))
                    {
                        int nlastdata = util.semestercount(bat, lastsemwithdata);
                        int nsem = nlastdata + 1;
                        string semx = util.incrementsemester(lastsemwithdata);
                        do
                        {
                            int retcol = retoffset + nsem;
                            sheet.Cells[nrow, bathd[semx] + 1].Formula = "=" + Cellname(nrow, bathd[semx])+ "*'" + retsheetname + "'!" + Cellname(prow[prog], retcol);
                            int isem = util.semestercount(bat, semx);
                            add_batsemref(nrow, bathd[semx] + 1, prog, bat, isem,batsheetname);
                            semx = util.incrementsemester(semx);
                            nsem++;

                        }
                        while (semx != util.incrementsemester(endsem) && nsem <= nsemtot);

                    }
                }
            }

            foreach (programclass pc in qprog)
            {
                string prog = pc.name;
                string nextsem = util.incrementsemester(lastsemwithdata);
                string semnewbatch = nextsem;
                do
                {
                    nrow++;
                    sheet.Cells[nrow, bathd["Program"] + 1] = prog;
                    sheet.Cells[nrow, bathd["Start"] + 1] = semnewbatch;
                    sheet.Cells[nrow, bathd[semnewbatch] + 1].Formula = "='" + mainsheetname + "'!" + Cellname(prow[prog], planhd[t1string + semnewbatch] + 1);
                    add_batsemref(nrow, bathd[semnewbatch] + 1, prog, semnewbatch, 1, batsheetname);
                    int nsem = 2;
                    string semx = util.incrementsemester(semnewbatch);
                    if (semnewbatch != endsem)
                    {
                        do
                        {
                            int retcol = retoffset + nsem;
                            sheet.Cells[nrow, bathd[semx] + 1].Formula = "=" + Cellname(nrow, bathd[semx]) + "*'" + retsheetname + "'!" + Cellname(prow[prog], retcol);
                            int isem = util.semestercount(semnewbatch, semx);
                            add_batsemref(nrow, bathd[semx] + 1, prog, semnewbatch, isem,batsheetname);
                            semx = util.incrementsemester(semx);
                            nsem++;

                        }
                        while (semx != util.incrementsemester(endsem) && nsem <= pc.semesters);
                    }
                    //sheet.Cells[nrow, bathd[applstring] + 1] = progbatappl[prog][bat];
                    //sheet.Cells[nrow, bathd[accstring] + 1] = progbatacc[prog][bat];
                    //int retcol = retoffset + nsem;
                    //sheet.Cells[nrow, bathd[semx] + 1].Formula = "=" + Cellname(nrow, bathd[semx]) + "*'" + retsheetname + "'!" + Cellname(prow[prog], retcol);
                    semnewbatch = util.incrementsemester(semnewbatch);
                    //nsem++;

                }
                while (semnewbatch != util.incrementsemester(endsem));
            }
            Excel.Range qa = sheet.Columns[1];
            qa.ColumnWidth = 50;
            sheet.Protect();
        }

        private void printfracprod(List<programclass> qprog, string inst)
        {
            memo("\t" + fracprodclass.printheader());
            foreach (programclass pc in qprog)
            {
                memo(pc.name + "\t" + pc.fracproddict[inst].print());
            }
        }

        private void Excelbutton_Click(object sender, EventArgs e)
        {
            lastsemwithdata = TBlastsem.Text;
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();

            string folder = util.timestampfolder(@"C:\Temp\Excel planning sheets per institution");
            //string folder = util.timestampfolder(Form1.folder + @"\Excel planning sheets per institution");
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);


            Dictionary<string, string> fninst = new Dictionary<string, string>();
            Dictionary<string, Excel.Workbook> xldict = new Dictionary<string, Excel.Workbook>();
            Dictionary<string, Dictionary<string, Excel.Worksheet>> sheetdictdict = new Dictionary<string, Dictionary<string, Excel.Worksheet>>();

            foreach (string inst in Form1.institutiondict.Keys)
            {
                fninst.Add(inst, util.unusedfn(folder + "HST-planering " + Form1.instshortdict[inst] +" "+util.yymmdd()+ ".xlsx"));
                Excel.Workbook xl = xlApp.Workbooks.Add();
                xldict.Add(inst, xl);
                sheetdictdict.Add(inst, new Dictionary<string, Excel.Worksheet>());
            }


            int ncat = 0;
            int maxcount = 333333;

            List<string> sheetnames = new List<string>();

            foreach (string inst in fninst.Keys)
            {
                memo(inst);
                if (inst != "Institutionen för information och teknik")
                    continue;

                List<programclass> qprog;
                if (RB_homeinst.Checked)
                    qprog = (from c in Form1.origprogramdict
                             where c.Value.utype != "Kurspaket"
                             where c.Value.homeinst == inst select c.Value).ToList();
                else
                    qprog = (from c in Form1.origprogramdict
                             where c.Value.utype != "Kurspaket"
                             where c.Value.fracproddict.ContainsKey(inst) select c.Value).ToList();

                var qpaket = (from c in Form1.origprogramdict
                         where c.Value.utype == "Kurspaket"
                         where c.Value.homeinst == inst
                         select c.Value).ToList();

                int nprog = qprog.Count;
                int allmaxsem = (from c in qprog select c.semesters).Max();

                switch (inst)
                {
                    case "Institutionen för information och teknik":
                        var q6 = (from c in qprog where c.semesters == 6 select c).ToList();
                        if (q6.Count > 0)
                        {
                            programclass p6 = programclass.clone(q6);
                            p6.name = "Nytt program 180 hp";
                            qprog.Add(p6);
                        }
                        var q4 = (from c in qprog where c.semesters == 4 where !c.is_advanced() select c).ToList();
                        if (q4.Count > 0)
                        {
                            programclass p4 = programclass.clone(q4);
                            p4.name = "Nytt program x-tekniker 120 hp";
                            qprog.Add(p4);
                        }
                        var qm = (from c in qprog where c.semesters == 4 where c.is_advanced() select c).ToList();
                        if (qm.Count > 0)
                        {
                            programclass pm = programclass.clone(qm);
                            pm.name = "Nytt masterprogram 120 hp";
                            qprog.Add(pm);
                        }
                        break;
                    default:
                        break;
                }

                printfracprod(qprog,inst);
                //foreach (string progname in Form1.programdict.Keys)
                //    if (Form1.programdict[progname].homeinst == inst)
                //        nprog++;
                memo("nprog = " + nprog);
                //private void AddExcelTab(Excel.Workbook xl, Dictionary<string, Excel.Worksheet> sheetdict, SortedDictionary<string, List<pubclass>> dict, string dictkey, List<string> sheetnames, int maxcount)
                //if (authorclass.instsubj.ContainsKey(auinst))
                //    foreach (string subj in authorclass.instsubj[auinst])
                //    {
                //        AddExcelTabDiva(xldict[auinst], sheetdictdict[auinst], ausubjpubdict, subj, sheetnames, maxcount, auinst);
                //    }
                //AddExcelTabDiva(xldict[auinst], sheetdictdict[auinst], auinstpubdict, auinst, sheetnames, maxcount, auinst);

                List<programclass> qcourse;
                qcourse = (from c in Form1.fkdict 
                           where c.Value.activecourse 
                           where c.Value.homeinst == inst 
                           select c.Value).OrderBy(c=>c.bestcode()).ToList();
                memo("# courses = " + qcourse.Count());

                string startsem = "VT22";
                string endsem = "HT26";
                
                Excel.Worksheet retsheet = xldict[inst].Sheets.Add();
                retsheet.Name = retsheetname;
                memo(retsheet.Name);
                sheetdictdict[inst].Add(retsheet.Name, retsheet);
                RetentionSheet(retsheet, qprog, allmaxsem);

                Excel.Worksheet mainsheet = xldict[inst].Sheets.Add();
                mainsheet.Name = mainsheetname;
                memo(mainsheet.Name);
                sheetdictdict[inst].Add(mainsheet.Name, mainsheet);

                Excel.Worksheet detailsheet = xldict[inst].Sheets.Add();
                detailsheet.Name = detailsheetname;
                memo(detailsheet.Name);
                sheetdictdict[inst].Add(detailsheet.Name, detailsheet);

                fill_planhd(mainsheet,detailsheet,qprog,startsem, endsem);

                Excel.Worksheet batsheet = xldict[inst].Sheets.Add();
                batsheet.Name = batsheetname;

                memo(batsheet.Name);
                sheetdictdict[inst].Add(batsheet.Name, batsheet);
                BatchSheet(batsheet, qprog, startsem, endsem);


                memo(mainsheet.Name);
                PlanSheet(mainsheet, qprog, startsem, endsem,inst);

                memo(detailsheet.Name);
                DetailSheet(detailsheet, qprog, startsem, endsem,inst);

                Excel.Worksheet coursesheet = xldict[inst].Sheets.Add();
                coursesheet.Name = coursesheetname;
                memo(coursesheet.Name);
                CourseSheet(coursesheet, qprog, qcourse, startsem, endsem);

                //mainsheet.Select();

                memo("Saving to " + fninst[inst]);
                xldict[inst].SaveAs(fninst[inst]);

                foreach (string sc in sheetdictdict[inst].Keys)
                {
                    Marshal.ReleaseComObject(sheetdictdict[inst][sc]);
                }
                xldict[inst].Close();
                Marshal.ReleaseComObject(xldict[inst]);


                //Excel.Worksheet sheet = (Excel.Worksheet)xldict[auinst].Sheets.Add();
                //sheet.Name = validsheetname(auinst, sheetnames);
                //sheetnames.Add(sheet.Name);
                //SheetWithHeader(sheet, auinstpubdict.Count);
                //auinstsheetdict.Add(auinst, sheet);
                //int publine = 1;
                //foreach (pubclass pc in auinstpubdict[auinst])
                //{
                //    publine++;
                //    pc.write_excelrow(sheet, publine, hd);
                //    if (publine > maxcount)
                //        break;
                //}
                //nauinst++;
                //if (nauinst > maxcount)
                //    break;
                break;
            }

            //memo("Saving to " + fnauinst);
            //xlWauinst.SaveAs(fnauinst);

            //Then you can read from the sheet, keeping in mind that indexing in Excel is not 0 based. This just reads the cells and prints them back just as they were in the file.

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            //for (int i = 1; i <= rowCount; i++)
            //{
            //    for (int j = 1; j <= colCount; j++)
            //    {
            //        //new line
            //        if (j == 1)
            //            Console.Write("\r\n");

            //        //write the value to the console
            //        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
            //            Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");

            //        //add useful things here!   
            //    }
            //}

            //Lastly, the references to the unmanaged memory must be released. If this is not properly done, then there will be lingering processes that will hold the file access writes to your Excel workbook.

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background

            //close and release

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            memo("==== DONE ====");

        }
    }
}
