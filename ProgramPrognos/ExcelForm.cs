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

        Dictionary<string, CheckBox> CBinst = new Dictionary<string, CheckBox>();

        bool prognos; 

        public ExcelForm()
        {
            int basey = 550;
            int pitch = 40;
            InitializeComponent();
            //foreach (string inst in Form1.institutiondict.Keys)
            //{
            //    CheckBox cb = new CheckBox();
            //    cb.Text = inst;
            //    cb.Location = new Point(800, basey);
            //    cb.AutoSize = true;
            //    cb.Visible = true;
            //    basey += pitch;
            //    this.Controls.Add(cb);
            //    CBinst.Add(inst, cb);
            //}
            //foreach (string inst in Form1.institutiondict.Keys)
            //{
            //    CBinst[inst].Visible = true;
            //}
            foreach (string inst in Form1.institutiondict.Keys)
            {
                LBinst.Items.Add(inst);
            }
            LBinst.Items.Add(Form1.hda);
            //this.Visible = true;
            //this.Refresh();
        }
        public void memo(string s)
        {
            richTextBox1.AppendText(s + "\n");
            richTextBox1.ScrollToCaret();
        }

        private void SheetWithHeader(Excel.Worksheet sheet, int datarows, Dictionary<string,int> hd)
        {
            if (sheet.Rows.Count < datarows)
            {
                for (int i = sheet.Rows.Count; i <= datarows; i++)
                    sheet.Rows.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);
            }
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
            rethd = new Dictionary<string, int>() { { "Program", 0 }, { "Medelretention", 1 }, { "Sökande -> T1", 2 }, { "U1 -> T1", 3 }, { "U2 -> T1", 4 }, { "Antagen -> T1", 5 } };
            int offset = retoffset;
            for (int i = 1; i < allmaxsem; i++)
            {
                rethd.Add("T" + i + "->T" + (i + 1), offset + i);
            }

            SheetWithHeader(retsheet, qprog.Count + 1, rethd);

            if (!rrow.ContainsKey(retsheet.Name))
            {
                //Dictionary<string, int> xrow;
                if (retsheet.Name.Contains("Kurser"))
                {
                    rrow.Add(retsheet.Name, ProgramNames(retsheet, qprog, courseoffset));
                }
                else
                {
                    rrow.Add(retsheet.Name, ProgramNames(retsheet, qprog, progoffset));
                }
            }
            //int offset = 3;
            int meancol = 2;

            foreach (programclass pc in qprog)
            {
                int row = rrow[retsheet.Name][pc.name];
                for (int j=0;j<4;j++)
                {
                    double trans = (pc.appltransition[j][1] == null) ? 0.8 : pc.appltransition[j][1].transitionprob;
                    retsheet.Cells[row, j + 3] = trans;
                }
                for (int i=0;i<=pc.semesters;i++)  //Malena 240423
                {
                    double trans = (pc.transition[i] == null) ? 0.8 : pc.transition[i].transitionprob;
                    retsheet.Cells[row, i + offset] = trans;
                }
                retsheet.Cells[row, meancol].Formula = "=AVERAGE(" + Cellname(row, offset+1) + ":" + Cellname(row, offset+1 + pc.semesters) + ")";

                if (row % 100 == 0)
                    memo(row.ToString());
            }
            retsheet.Range["B2", Cellname(qprog.Count + 2, allmaxsem + 3)].NumberFormat = "###.0%";
            retsheet.Cells[1, 1].Locked = false;
            retsheet.Protect();
        }

        //Dictionary<string, int> prow;
        //Dictionary<string, int> crow;
        Dictionary<string, Dictionary<string, int>> rrow = new Dictionary<string, Dictionary<string, int>>();
        Dictionary<string, int> planhd;
        Dictionary<string, int> plan2hd;
        Dictionary<string, int> bathd;
        Dictionary<string, int> rethd;
        Dictionary<string, int> coursehd;
        Dictionary<string, int> sumhd;
        Dictionary<string, int> triangelhd;
        Dictionary<string, int> progkurshd;
        Dictionary<string, int> fksumrow;
        int retoffset = 5;
        int courseoffset = 2;
        int progoffset = 2;
        int progkursoffset = 3;
        int sumoffset = 6;
        int triangeloffset = 8;
        string acceptstring = "Antas slut ";
        string acceptu1string = "Antas U1 ";
        string acceptu2string = "Antas U2 ";
        string t1string = "T1 ";
        string studstring = "Stud ";
        string fkstudstring = "FK-stud ";
        string fkstring = "FK ";
        string progstudstring = "Prog-stud ";
        string progstring = "prog ";
        string hststring = "HST ";
        string hprstring = "HPR ";
        string moneystring = "Kr ";
        string instsumstring = "Summa inst ";
        string retsheetname = "RetentionProgram";
        string retcoursesheetname = "RetentionKurser";
        string retpaketsheetname = "RetentionPaket";
        string mainsheetname = "Planering";
        string detailsheetname = "Detaljer";
        string batsheetname = "Programkullar";
        string coursesheetname = "Kurser";
        string paysheetname = "Betalande stud";
        string paketsheetname = "Kurspaket";
        string sumsheetname = "Summa";
        string triangelsheetname = "Trianglar";
        string progkurssheetname = "Prog-kurs";
        Dictionary<int, string> progkurssheetnames = new Dictionary<int, string>();
        string applstring = "1:ahand ";
        string accstring = "Antagna";
        string studhststring = "Stud/HST?";
        string prestationstring = "Prest-grad";
        string budgetstring = "Budget plan-tal ";
        string diffstring = "Diff prognos-budget ";
        string inststring = "Institution";
        string newcoursename = "Ny kurs ";

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

            SheetWithHeader(sheet, qprog.Count + 6, planhd);
            SheetWithHeader(sheet2, qprog.Count + 2, plan2hd);

            //sheet.Protect();

        }

        private void fill_planhd_prognos(Excel.Worksheet sheet, Excel.Worksheet sheet2, List<programclass> qprog, string startsem, string endsem)
        {
            planhd = new Dictionary<string, int>() { { "Program", 0 }, { inststring, 1 } };
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
                planhd.Add(applstring + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));

            sem = startsem;
            //col = planhd.Count + 1;
            do
            {
                planhd.Add(acceptu1string + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));

            sem = startsem;
            //col = planhd.Count + 1;
            do
            {
                planhd.Add(acceptu2string + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));

            sem = startsem;
            //col = planhd.Count + 1;
            do
            {
                planhd.Add(acceptstring + sem, col);
                col++;
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
            sem = endsem;
            planhd.Add(budgetstring + sem, col);
            col++;
            planhd.Add(diffstring + sem, col);
            col++;

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

            SheetWithHeader(sheet, qprog.Count + 6, planhd);
            SheetWithHeader(sheet2, qprog.Count + 2, plan2hd);

            //sheet.Protect();

        }

        private void PlanSheet(Excel.Worksheet sheet, List<programclass> qprog, string startsem, string endsem,string inst)
        {
            if (!rrow.ContainsKey(sheet.Name))
                rrow.Add(sheet.Name, ProgramNames(sheet, qprog, progoffset));
            Dictionary<string, bool> phtstart = new Dictionary<string, bool>();
            Dictionary<string, bool> pvtstart = new Dictionary<string, bool>();
            sheet.Cells[2, 1] = "Total";
            sheet.Rows[2].Font.Bold = true;
            for (int icol = 2; icol <= planhd.Count; icol++)
            {
                sheet.Cells[2, icol].Formula = "=SUM(" + Cellname(3, icol) + ":" + Cellname(3 + qprog.Count+1, icol) + ")";
            }

            sheet.Range["A1", Cellname(1,planhd.Count)].WrapText = true;



            //='Sheet 1'!A3

            int startyear = util.semtoint(startsem);
            int endyear = util.semtoint(endsem);

            double roundfactor = 5;

            int lastcolwithdata = -1;

            foreach (programclass pc in qprog)
            {
                int row = rrow[sheet.Name][pc.name];
                for (int year = startyear; year <= endyear; year++)
                {
                    int year4 = year + 2000;
                    int colyear = planhd[moneystring + "20" + year] + 1;
                    double hstpeng = pc.fracproddict[inst].hstpeng;
                    if (hstpeng == 0)
                        hstpeng = qprog.First().fracproddict[inst].hstpeng;
                    double hprpeng = pc.fracproddict[inst].hprpeng;
                    if (hprpeng == 0)
                        hprpeng = qprog.First().fracproddict[inst].hprpeng;
                    hstpeng *= Form1.get_pengindex(year);
                    hprpeng *= Form1.get_pengindex(year);
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
                        if (pc.name.StartsWith("Nytt ") || pc.coursedict.Count == 0)
                        {
                            // OLD ALGORITHM, used for new programs:
                            sheet.Cells[row, colyear].Formula = toreplace + "=" + hstpeng + "*'" + detailsheetname + "'!" + Cellname(row, colhst) + "+" + hprpeng + "*'" + detailsheetname + "'!" + Cellname(row, colhpr);
                        }
                        else
                        {
                            // NEW ALGORITHM, used for old programs:
                            sheet.Cells[row, colyear].Formula = toreplace + "='" + progkurssheetnames[year4] + "'!" + Cellname(2, progkurshd[pc.name] + 1);
                        }
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
                    if (sem == lastsemwithdata)
                        lastcolwithdata = col;
                    if (pc.fk)
                    {
                        if (!pc.name.Contains("från mediaområdet")) //blir dubbelräknat om de tas med
                        {
                            //double hst = pc.totalprod.frachst;
                            //double hst = pc.fracproddict[inst].frachst*pc.totalprod.frachst;
                            string uo = Form1.utbomrdict[pc.name];
                            int year = util.yearfromsem(sem);
                            string hd = hststring + fkstring + uo + " " + year;
                            string f = toreplace + "=0,5*'" + coursesheetname + "'!" + Cellname(2, coursehd[hd] + 1);
                            sheet.Cells[row, col].Formula = f;
                            phtstart[pc.name] = true;
                            pvtstart[pc.name] = true;
                        }
                    }
                    else
                    {
                        programbatchclass bc = (from c in pc.batchlist where c.batchstart == sem select c).FirstOrDefault();
                        if (bc != null)
                        {
                            if (bc.getstud(1) > 0)
                                sheet.Cells[row, col] = bc.getstud(1);
                            if (bc.batchstart.Contains("H"))
                                phtstart[pc.name] = true;
                            else
                                pvtstart[pc.name] = true;
                        }
                    }
                    sem = util.incrementsemester(sem);
                }
                while (sem != util.shiftsemester(lastsemwithdata,14));

                if (specialinputdict.ContainsKey(pc.name))
                {
                    foreach (string h in specialinputdict[pc.name].Keys)
                    {
                        if (planhd.ContainsKey(h))
                        {
                            sheet.Cells[row, planhd[h] + 1] = specialinputdict[pc.name][h];
                        }
                    }
                }

            }
            //sheet.Range["B2", Cellname(qprog.Count + 1, allmaxsem + 2)].NumberFormat = "###.0%";
            //sheet.Cells[1, 1].Locked = false;
            int endkrcol = endyear - startyear + 2;
            sheet.Range["B2", Cellname(qprog.Count + 7, endkrcol)].NumberFormat = "### ### ###";
            sheet.Range["B2", Cellname(qprog.Count + 3, endkrcol)].Interior.Color = Excel.XlRgbColor.rgbLightPink;
            for (int i = 2; i <= endkrcol; i++)
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
                int row = rrow[sheet.Name][pc.name];
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
        }

        private void PlanSheetPrognos(Excel.Worksheet sheet, List<programclass> qprog, string startsem, string endsem, string inst)
        {
            if (!rrow.ContainsKey(sheet.Name))
                rrow.Add(sheet.Name, ProgramNames(sheet, qprog, progoffset));
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

            string prognossem = util.incrementsemester(lastsemwithdata);
            memo("prognossem = " + prognossem);

            double roundfactor = 5;

            int lastcolwithdata = -1;

            foreach (programclass pc in qprog)
            {
                int row = rrow[sheet.Name][pc.name];
                int detailrow = rrow[detailsheetname][pc.name];
                for (int year = startyear; year <= endyear; year++)
                {
                    int colyear = planhd[moneystring + "20" + year] + 1;
                    double hstpeng = pc.fracproddict.ContainsKey(inst)? pc.fracproddict[inst].hstpeng : 0;
                    if (hstpeng == 0)
                        hstpeng = qprog.First().fracproddict[inst].hstpeng;
                    double hprpeng = pc.fracproddict.ContainsKey(inst) ? pc.fracproddict[inst].hprpeng : 0;
                    if (hprpeng == 0)
                        hprpeng = qprog.First().fracproddict[inst].hprpeng;
                    hstpeng *= Form1.get_pengindex(year);
                    hprpeng *= Form1.get_pengindex(year);
                    double frachst = 1;
                    double frachpr = 1;
                    if (inst != Form1.hda && pc.fracproddict.ContainsKey(inst) && pc.fracproddict[Form1.hda].frachst > 0)
                    {
                        frachst = pc.fracproddict[inst].frachst/ pc.fracproddict[Form1.hda].frachst;
                        frachpr = pc.fracproddict[inst].frachpr/ pc.fracproddict[Form1.hda].frachpr;
                    }
                    if (pc.fk)
                    {
                        int colvt = planhd[t1string + "VT" + year] + 1;
                        int colht = planhd[t1string + "HT" + year] + 1;
                        //sheet.Cells[row, colyear].Formula = toreplace + "=" + (hstpeng*frachst + hprpeng*frachpr) + "*(" + Cellname(row, colvt) + "+" + Cellname(row, colht) + ")";
                        sheet.Cells[row, colyear].Formula = toreplace + "=" + (hstpeng + hprpeng) + "*(" + Cellname(row, colvt) + "+" + Cellname(row, colht) + ")";
                    }
                    else
                    {
                        int colhst = plan2hd[hststring + "20" + year] + 1;
                        int colhpr = plan2hd[hprstring + "20" + year] + 1;
                        sheet.Cells[row, colyear].Formula = toreplace + "=" + hstpeng*frachst + "*'" + detailsheetname + "'!" + Cellname(detailrow, colhst) + "+" + hprpeng*frachpr + "*'" + detailsheetname + "'!" + Cellname(detailrow, colhpr);
                    }
                }

                string studhst = pc.fk ? "HST" : "Stud";
                sheet.Cells[row, planhd[studhststring] + 1] = studhst;

                phtstart.Add(pc.name, false);
                pvtstart.Add(pc.name, false);

                string sem = startsem;
                do
                {
                    int acccol = planhd[acceptstring + sem] + 1;
                    int u1col = planhd[acceptu1string + sem] + 1;
                    int u2col = planhd[acceptu2string + sem] + 1;
                    int applcol = planhd[applstring + sem] + 1;

                    int col = planhd[t1string + sem] + 1;
                    lastcolwithdata = col;
                    if (pc.fk)
                    {
                        //double hst = pc.totalprod.frachst;
                        double hst = 0.5*pc.fracproddict[inst].frachst * pc.totalprod.frachst;
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

                            if (bc.applicants[0] != null)
                                sheet.Cells[row, applcol] = (double)bc.applicants[0];
                            if (bc.applicants[1] != null)
                                sheet.Cells[row, u1col] = (double)bc.applicants[1];
                            if (bc.applicants[2] != null)
                                sheet.Cells[row, u2col] = (double)bc.applicants[2];
                            if (bc.applicants[3] != null)
                                sheet.Cells[row, acccol] = (double)bc.applicants[3];
                        }

                    }
                    sem = util.incrementsemester(sem);
                }
                while (sem != util.incrementsemester(lastsemwithdata));

                sem = prognossem;
                if (!pc.fk)
                {
                    programbatchclass bc = (from c in pc.batchlist where c.batchstart == sem select c).FirstOrDefault();

                    int col = planhd["T1 " + sem]+1;
                    int acccol = planhd[acceptstring + sem] + 1;
                    int u1col = planhd[acceptu1string + sem] + 1;
                    int u2col = planhd[acceptu2string + sem] + 1;
                    int applcol = planhd[applstring + sem] + 1;
                    int retacccol = rethd["Antagen -> T1"] + 1;
                    int retu1col = rethd["U1 -> T1"] + 1;
                    int retu2col = rethd["U2 -> T1"] + 1;
                    int retapplcol = rethd["Sökande -> T1"] + 1;

                    if (bc != null && bc.getstud(1) > 0)
                    {
                        sheet.Cells[row, col] = bc.getstud(1);
                    }
                    else
                    {
                        string f = toreplace + "=IF(" + Cellname(row, acccol) + ">0;" + retsheetname + "!" + Cellname(row, retacccol) + "*" + Cellname(row, acccol) + ";"
                            + "IF(" + Cellname(row, u2col) + " > 0; " + retsheetname + "!" + Cellname(row, retu2col) + "*" + Cellname(row, u2col) + "; "
                            + "IF(" + Cellname(row, u1col) + " > 0; " + retsheetname + "!" + Cellname(row, retu1col) + "*" + Cellname(row, u1col) + "; "
                            + "IF(" + Cellname(row, applcol) + " > 0; " + retsheetname + "!" + Cellname(row, retapplcol) + "*" + Cellname(row, applcol) + ";0))))";
                        sheet.Cells[row, col].Formula = f;
                    }
                    //programbatchclass bc = (from c in pc.batchlist where c.batchstart == sem select c).FirstOrDefault();
                    if (bc != null)
                    {
                        if (bc.applicants[0] != null)
                            sheet.Cells[row, applcol] = (double)bc.applicants[0];
                        if (bc.applicants[1] != null)
                            sheet.Cells[row, u1col] = (double)bc.applicants[1];
                        if (bc.applicants[2] != null)
                            sheet.Cells[row, u2col] = (double)bc.applicants[2];
                        if (bc.applicants[3] != null)
                            sheet.Cells[row, acccol] = (double)bc.applicants[3];
                    }

                    int budgetcol = planhd[budgetstring + sem] + 1;
                    if (bc != null)
                        sheet.Cells[row, budgetcol] = bc.budget_T1;
                    else
                        sheet.Cells[row, budgetcol] = 0;

                    int diffcol = planhd[diffstring + sem] + 1;
                    sheet.Cells[row, diffcol].Formula = toreplace + "=" + Cellname(row, col) + "-" + Cellname(row, budgetcol);

                }
                //sem = util.incrementsemester(sem);

            }
            //sheet.Range["B2", Cellname(qprog.Count + 1, allmaxsem + 2)].NumberFormat = "###.0%";
            //sheet.Cells[1, 1].Locked = false;
            sheet.Range["B2", Cellname(qprog.Count + 6, 4)].NumberFormat = "### ### ###";
            sheet.Range["B2", Cellname(qprog.Count + 2, 3)].Interior.Color = Excel.XlRgbColor.rgbLightPink;
            sheet.Range["D2", Cellname(qprog.Count + 2, 4)].Interior.Color = Excel.XlRgbColor.rgbPink;
            
            Excel.Range qainst = sheet.Columns[2];
            qainst.ColumnWidth = 6;
            for (int i = 3; i < 5; i++)
            {
                Excel.Range qa = sheet.Columns[i];
                qa.ColumnWidth = 15;
            }

            for (int i = 5; i < 17; i++)
            {
                Excel.Range qa = sheet.Columns[i];
                qa.ColumnWidth = 12;
            }

            for (int i = 17; i < 21; i++)
            {
                Excel.Range qa = sheet.Columns[i];
                qa.ColumnWidth = 13.5;
            }

            //for (int i = lastcolwithdata+1; i <=planhd.Count; i++)
            //{
            //    Excel.Range qa = sheet.Columns[i];
            //    //qa.ColumnWidth = 15;
            //    qa.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            //}

            PrognosColors(sheet, qprog,lastcolwithdata,prognossem);

            //sheet.Range[Cellname(3, lastcolwithdata + 1), Cellname(qprog.Count + 2, planhd.Count + 1)].Locked = false;
            //sheet.Range[Cellname(3, lastcolwithdata + 1), Cellname(qprog.Count + 2, planhd.Count)].Interior.Color = Excel.XlRgbColor.rgbPaleGoldenrod;
            //sheet.Range[Cellname(3, lastcolwithdata + 1), Cellname(qprog.Count + 2, planhd.Count)].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            //var qadiff = sheet.Range[Cellname(2, planhd[diffstring + prognossem] + 1), Cellname(qprog.Count + 2, planhd[diffstring + prognossem])];
            //Excel.FormatCondition cond = qadiff.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlLess, 0);
            //cond.Font.Color = Color.Red;
            //qadiff.NumberFormat = "###";


            foreach (programclass pc in qprog)
            {
                int row = rrow[sheet.Name][pc.name];
                bool ht = phtstart[pc.name];
                bool vt = pvtstart[pc.name];
                int icol = lastcolwithdata + 1;
                while (icol <= planhd.Count)
                {
                    if (ht)
                        sheet.Cells[row, icol].Interior.Color = Excel.XlRgbColor.rgbYellow;
                    if (vt && icol < planhd.Count)
                        sheet.Cells[row, icol + 1].Interior.Color = Excel.XlRgbColor.rgbYellow;
                    icol += 2;
                }

            }

            //int acccol2 = planhd[acceptstring + prognossem]+1;
            //int u1col2 = planhd[acceptu1string + prognossem]+1;
            //int u2col2 = planhd[acceptu2string + prognossem]+1;
            //int applcol2 = planhd[applstring + prognossem]+1;
            //int prognoscol = planhd[t1string + prognossem]+1;
            //int budgetcol2 = planhd[budgetstring + prognossem] + 1;
            //int diffcol2 = planhd[diffstring + prognossem] + 1;

            //sheet.Range[Cellname(3, acccol2), Cellname(qprog.Count + 2, acccol2)].Interior.Color = Excel.XlRgbColor.rgbGreen;
            //sheet.Range[Cellname(3, u1col2), Cellname(qprog.Count + 2,u1col2)].Interior.Color = Excel.XlRgbColor.rgbLawnGreen;
            //sheet.Range[Cellname(3, u2col2), Cellname(qprog.Count + 2, u2col2)].Interior.Color = Excel.XlRgbColor.rgbLightGreen;
            //sheet.Range[Cellname(3, applcol2), Cellname(qprog.Count + 2, applcol2)].Interior.Color = Excel.XlRgbColor.rgbMediumSpringGreen;
            //sheet.Range[Cellname(3, prognoscol), Cellname(qprog.Count + 2, prognoscol)].NumberFormat = "###";
            //sheet.Range[Cellname(3, budgetcol2), Cellname(qprog.Count + 2, budgetcol2)].Interior.Color = Excel.XlRgbColor.rgbGold;
            //sheet.Range[Cellname(3, diffcol2), Cellname(qprog.Count + 2, diffcol2)].Interior.Color = Excel.XlRgbColor.rgbLime;


            //sheet.FreezeColumns(1);
        }

        private void PrognosColors(Excel.Worksheet sheet, List<programclass> qprog, int lastcolwithdata, string prognossem)
        {
            sheet.Range["B2", Cellname(qprog.Count + 6, 4)].NumberFormat = "### ### ###";
            sheet.Range["B2", Cellname(qprog.Count + 2, 4)].Interior.Color = Excel.XlRgbColor.rgbLightPink;
            for (int i = 2; i < 4; i++)
            {
                Excel.Range qa = sheet.Columns[i];
                qa.ColumnWidth = 15;
            }

            for (int i = 5; i < 17; i++)
            {
                Excel.Range qa = sheet.Columns[i];
                qa.ColumnWidth = 12;
            }

            for (int i = 17; i < 21; i++)
            {
                Excel.Range qa = sheet.Columns[i];
                qa.ColumnWidth = 13.5;
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

            var qadiff = sheet.Range[Cellname(2, planhd[diffstring + prognossem] + 1), Cellname(qprog.Count + 2, planhd[diffstring + prognossem])];
            Excel.FormatCondition cond = qadiff.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlLess, 0);
            cond.Font.Color = Color.Red;
            qadiff.NumberFormat = "###";

            int acccol2 = planhd[acceptstring + prognossem] + 1;
            int u1col2 = planhd[acceptu1string + prognossem] + 1;
            int u2col2 = planhd[acceptu2string + prognossem] + 1;
            int applcol2 = planhd[applstring + prognossem] + 1;
            int prognoscol = planhd[t1string + prognossem] + 1;
            int budgetcol2 = planhd[budgetstring + prognossem] + 1;
            int diffcol2 = planhd[diffstring + prognossem] + 1;

            sheet.Range[Cellname(3, acccol2), Cellname(qprog.Count + 2, acccol2)].Interior.Color = Excel.XlRgbColor.rgbGreen;
            sheet.Range[Cellname(3, u1col2), Cellname(qprog.Count + 2, u1col2)].Interior.Color = Excel.XlRgbColor.rgbLawnGreen;
            sheet.Range[Cellname(3, u2col2), Cellname(qprog.Count + 2, u2col2)].Interior.Color = Excel.XlRgbColor.rgbLightGreen;
            sheet.Range[Cellname(3, applcol2), Cellname(qprog.Count + 2, applcol2)].Interior.Color = Excel.XlRgbColor.rgbMediumSpringGreen;
            sheet.Range[Cellname(3, prognoscol), Cellname(qprog.Count + 2, prognoscol)].NumberFormat = "###";
            sheet.Range[Cellname(3, budgetcol2), Cellname(qprog.Count + 2, budgetcol2)].Interior.Color = Excel.XlRgbColor.rgbGold;
            sheet.Range[Cellname(3, diffcol2), Cellname(qprog.Count + 2, diffcol2)].Interior.Color = Excel.XlRgbColor.rgbLime;


        }

        private void PrognosColorsCourse(Excel.Worksheet sheet, List<programclass> qprog, int lastcolwithdata, string prognossem)
        {
            sheet.Range["G2", Cellname(qprog.Count + 6, 10)].NumberFormat = "### ### ###";
            sheet.Range["G2", Cellname(qprog.Count + 2, 10)].Interior.Color = Excel.XlRgbColor.rgbLightPink;
            for (int i = 2; i < 4; i++)
            {
                Excel.Range qa = sheet.Columns[i];
                qa.ColumnWidth = 15;
            }

            for (int i = 5; i < 17; i++)
            {
                Excel.Range qa = sheet.Columns[i];
                qa.ColumnWidth = 12;
            }

            for (int i = 17; i < 21; i++)
            {
                Excel.Range qa = sheet.Columns[i];
                qa.ColumnWidth = 13.5;
            }

            //for (int i = lastcolwithdata+1; i <=coursehd.Count; i++)
            //{
            //    Excel.Range qa = sheet.Columns[i];
            //    //qa.ColumnWidth = 15;
            //    qa.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            //}

            sheet.Range[Cellname(3, lastcolwithdata + 1), Cellname(qprog.Count + 2, coursehd.Count + 1)].Locked = false;
            sheet.Range[Cellname(3, lastcolwithdata + 1), Cellname(qprog.Count + 2, coursehd.Count)].Interior.Color = Excel.XlRgbColor.rgbPaleGoldenrod;
            sheet.Range[Cellname(3, lastcolwithdata + 1), Cellname(qprog.Count + 2, coursehd.Count)].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            var qadiff = sheet.Range[Cellname(2, coursehd[diffstring + prognossem] + 1), Cellname(qprog.Count + 2, coursehd[diffstring + prognossem])];
            Excel.FormatCondition cond = qadiff.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlLess, 0);
            cond.Font.Color = Color.Red;
            qadiff.NumberFormat = "###";


            foreach (string s in coursehd.Keys)
            {
                int ncol = coursehd[s] + 1;
                Color color = Color.White;
                if (s.Contains(moneystring + fkstring))
                    color = Color.Pink;
                else if (s.Contains(moneystring + progstring))
                    color = Color.LightPink;
                //else if (s.Contains(fkstudstring))
                //    color = Color.Yellow;
                else if (s.Contains(progstudstring))
                    color = Color.Tan;
                //else if (s.Contains(acceptstring))
                //    color = Color.LightGreen;
                else if (s.Contains(hststring))
                    color = Color.LightBlue;
                else if (s.Contains(hprstring))
                    color = Color.PaleTurquoise;
                else if (s.Contains(studstring))
                    color = Color.AntiqueWhite;
                sheet.Range[Cellname(3, ncol), Cellname(qprog.Count+2, ncol)].Interior.Color = color;

                var qa = sheet.Range[Cellname(3, ncol), Cellname(qprog.Count+2, ncol)];
                if (color == Color.Yellow)
                {
                    qa.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    qa.Locked = false;
                }
                if (color == Color.AntiqueWhite)
                {
                    Excel.FormatCondition cond2 = qa.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlBetween, 0.01, "=$c$2");
                    cond2.Font.Color = Color.Red;
                    qa.NumberFormat = "###";
                }
            }

            int acccol2 = coursehd[acceptstring + prognossem] + 1;
            int u1col2 = coursehd[acceptu1string + prognossem] + 1;
            int u2col2 = coursehd[acceptu2string + prognossem] + 1;
            int applcol2 = coursehd[applstring + prognossem] + 1;
            int prognoscol = coursehd[fkstudstring + prognossem] + 1;
            int budgetcol2 = coursehd[budgetstring + prognossem] + 1;
            int diffcol2 = coursehd[diffstring + prognossem] + 1;

            sheet.Range[Cellname(3, acccol2), Cellname(qprog.Count + 2, acccol2)].Interior.Color = Excel.XlRgbColor.rgbGreen;
            sheet.Range[Cellname(3, u1col2), Cellname(qprog.Count + 2, u1col2)].Interior.Color = Excel.XlRgbColor.rgbLawnGreen;
            sheet.Range[Cellname(3, u2col2), Cellname(qprog.Count + 2, u2col2)].Interior.Color = Excel.XlRgbColor.rgbLightGreen;
            sheet.Range[Cellname(3, applcol2), Cellname(qprog.Count + 2, applcol2)].Interior.Color = Excel.XlRgbColor.rgbMediumSpringGreen;
            sheet.Range[Cellname(3, prognoscol), Cellname(qprog.Count + 2, prognoscol)].NumberFormat = "###";
            sheet.Range[Cellname(3, prognoscol), Cellname(qprog.Count + 2, prognoscol)].Interior.Color = Excel.XlRgbColor.rgbYellow;
            sheet.Range[Cellname(3, budgetcol2), Cellname(qprog.Count + 2, budgetcol2)].Interior.Color = Excel.XlRgbColor.rgbGold;
            sheet.Range[Cellname(3, diffcol2), Cellname(qprog.Count + 2, diffcol2)].Interior.Color = Excel.XlRgbColor.rgbLime;


        }

        // batsemref[prog][batstart][sem] = Cellref.

        private void PlanSheet_FKrows(Excel.Worksheet sheet, List<programclass> qprog, string startsem, string endsem, string inst)
        {

            int startyear = util.semtoint(startsem);
            int endyear = util.semtoint(endsem);

            int progendrow = qprog.Count() + 2;
            int totalrow = 2;
            if (!rrow.ContainsKey(sheet.Name))
                rrow.Add(sheet.Name, ProgramNames(sheet, qprog, progoffset));

            int paketrow = 1;
            int instructionrow = 3;
            int kursfromflikrow = 4;
            int kursfrompinkrow = 5;
            sheet.Cells[progendrow + paketrow, 1].Value = "Summa kurspaket från fliken 'Kurspaket'";
            sheet.Cells[progendrow + kursfromflikrow, 1].Value = "Summa FK från fliken 'Kurser'";
            sheet.Cells[progendrow + instructionrow, 1].Value = "'-- Nedanstående två rader ska bli samma summa --";
            //sheet.Cells[progendrow + 5, 1].Value = "Summa FK+kurspaket från flikar";
            sheet.Cells[progendrow + kursfrompinkrow, 1].Value = "Summa FK i rosa fält ovan";

            for (int year = startyear; year <= endyear; year++)
            {
                int colyear = planhd[moneystring + "20" + year] + 1;
                int coursecolyear = coursehd[moneystring + fkstring + "20" + year] + 1;

                int fsrow = totalrow;
                if (inst != Form1.hda & prognos)
                    fsrow = fksumrow[inst];
                sheet.Cells[progendrow + kursfromflikrow, colyear].Formula = toreplace + "='" + coursesheetname + "'!" + Cellname(fsrow, coursecolyear);
                if (inst == Form1.hda || !prognos)
                    sheet.Cells[progendrow + paketrow, colyear].Formula = toreplace + "='" + paketsheetname + "'!" + Cellname(totalrow, colyear);

                //sheet.Cells[progendrow + 5, colyear].Formula = toreplace + "="+Cellname(progendrow+2,colyear)+"+"+Cellname(progendrow+3,colyear);
                string fklines = toreplace + "=";
                foreach (programclass pc in qprog)
                {
                    if (pc.fk)
                    {
                        fklines += Cellname(rrow[sheet.Name][pc.name], colyear)+"+";
                    }
                }
                fklines = fklines.Trim('+');
                sheet.Cells[progendrow + kursfrompinkrow, colyear].Formula = fklines;

            }

            //sheet.Protect();



        }

        private void ProgkursSheet(Excel.Worksheet sheet, List<programclass> qprog, List<programclass> qcourse, int year, string inst)
        {
            progkurshd = new Dictionary<string, int>() { { "Kurs", 0 }, { "Kurskod", 1 }, { "Ämne", 2 } };

            int jcol = progkurshd.Count;
            foreach (programclass pc in qprog)
            {
                progkurshd.Add(pc.name, jcol);
                jcol++;
            }

            SheetWithHeader(sheet, qprog.Count + progkursoffset, progkurshd);

            if (!rrow.ContainsKey(sheet.Name))
                rrow.Add(sheet.Name, ProgramNames(sheet, qcourse, progkursoffset));
            sheet.Cells[2, 1] = "Total";
            sheet.Rows[2].Font.Bold = true;
            for (int icol = 4; icol <= progkurshd.Count; icol++)
            {
                sheet.Cells[2, icol].Formula = "=SUM(" + Cellname(3, icol) + ":" + Cellname(3 + qcourse.Count, icol) + ")";
            }

            string vt = "VT" + year % 100;
            string ht = "HT" + year % 100;
            int vtcol = bathd[vt] + 1;
            int htcol = bathd[ht] + 1;

            sheet.Cells[3, 1] = "Stud";
            sheet.Rows[3].Font.Bold = true;
            for (int icol = 4; icol <= progkurshd.Count; icol++)
            {
                sheet.Cells[3, icol].Formula = toreplace+"=SUMIF("+batsheetname+"!A2:"+Cellname(batsheetrow,htcol)+";"+Cellname(1,icol)+";"+batsheetname+"!"+Cellname(2,vtcol)+":"+Cellname(batsheetrow,htcol)+")";
            }

            //='Sheet 1'!A3
            double krhstsum = 0;
            double prestsum = 0;
            int ncourses = 0;

            Dictionary<string, int> coderow = new Dictionary<string, int>();
            Dictionary<string, double> krstuddict = new Dictionary<string, double>();

            foreach (programclass cc in qcourse)
            {
                if (cc.homeinst == Form1.utaninst)
                    continue;
                int row = rrow[sheet.Name][cc.name];

                sheet.Cells[row, 1] = cc.name;
                sheet.Cells[row, 2] = cc.bestcode();
                sheet.Cells[row, 3] = cc.subjectcode;

                bool nykurs = cc.name.StartsWith(newcoursename);
                double prest = 0.8;
                if (cc.totalprod.frachst > 0)
                {
                    prest = cc.totalprod.frachpr / cc.totalprod.frachst;
                    if (prest > 1)
                        prest = 1;
                }
                else if (nykurs)
                {
                    prest = prestsum / ncourses;
                }
                prestsum += prest;

                double krhst;
                if (nykurs)
                    krhst = krhstsum / ncourses;
                else
                    krhst = Form1.hstkr(1, cc.studentpengarea, -1) + prest * Form1.hprkr(1, cc.studentpengarea, -1);
                //coursesheet.Cells[nrow, moneycol].Value = krhst;
                krhstsum += krhst;
                ncourses++;
                double hpfix = cc.hp > 0 ? cc.hp : 7.5;
                double krstud = krhst * hpfix / 60 * Form1.get_pengindex(year);
                krstuddict.Add(cc.bestcode(), krstud);

                coderow.Add(cc.bestcode(), row);
            }

            foreach (programclass pc in qprog)
            {
                int icol = progkurshd[pc.name] + 1;
                foreach (string sm in new List<string>(){vt,ht })
                {
                    foreach (int isem in pc.coursedict.Keys)
                    {
                        string batstart = util.find_batstart(sm, isem);
                        if (batsemref[pc.name].ContainsKey(batstart)
                            && batsemref[pc.name][batstart].ContainsKey(isem))
                        {
                            foreach (string ccode in pc.coursedict[isem].Keys)
                            {
                                if (!coderow.ContainsKey(ccode))
                                    continue;
                                string sumref = "";
                                if (sheet.Cells[coderow[ccode], icol].Value != null)
                                    sumref = sheet.Cells[coderow[ccode], icol].Value;
                                if (String.IsNullOrEmpty(sumref))
                                    sumref = toreplace + "=";
                                else
                                    sumref += "+";
                                sumref += batsemref[pc.name][batstart][isem];
                                sumref += "*" + pc.coursedict[isem][ccode];
                                sumref += "*" + krstuddict[ccode];
                                sheet.Cells[coderow[ccode], icol].Value = sumref;
                            }
                        }
                    }
                }

            }


            //sheet.Range["B2", Cellname(qprog.Count + 1, allmaxsem + 2)].NumberFormat = "###.0%";
            //sheet.Cells[1, 1].Locked = false;
            sheet.Range["d2", Cellname(qprog.Count + 2, progkurshd.Count)].NumberFormat = "### ### ###";
            //sheet.Range["B2", Cellname(qprog.Count + 2, 6)].Interior.Color = Excel.XlRgbColor.rgbLightPink;
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

            //sheet.Range[Cellname(3, lastcolwithdata + 1), Cellname(qprog.Count + 2, planhd.Count + 1)].Locked = false;
            //sheet.Range[Cellname(3, lastcolwithdata + 1), Cellname(qprog.Count + 2, planhd.Count)].Interior.Color = Excel.XlRgbColor.rgbPaleGoldenrod;
            //sheet.Range[Cellname(3, lastcolwithdata + 1), Cellname(qprog.Count + 2, planhd.Count)].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;


            //sheet.FreezeColumns(1);
            //sheet.Protect();
        }


        private void PaketSheet(Excel.Worksheet sheet, List<programclass> qprog, string startsem, string endsem, string inst)
        {
            SheetWithHeader(sheet, qprog.Count + 2, planhd);

            if (!rrow.ContainsKey(sheet.Name))
                rrow.Add(sheet.Name, ProgramNames(sheet, qprog, progoffset));
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
                if (pc.homeinst == Form1.utaninst)// && !(inst==Form1.hda))
                    continue;
                int row = rrow[sheet.Name][pc.name];
                for (int year = startyear; year <= endyear; year++)
                {
                    int colyear = planhd[moneystring + "20" + year] + 1;
                    string firstcourse = pc.coursedict[1].Keys.First();
                    programclass cc = Form1.fkcodedict[firstcourse];
                    double hstpeng = Form1.hstkr(1, cc.studentpengarea,year);
                    double hprpeng = Form1.hprkr(1, cc.studentpengarea,year);
                    //double hstpeng = pc.fracproddict[inst].hstpeng;
                    //if (hstpeng == 0)
                    //    hstpeng = qprog.First().fracproddict[inst].hstpeng;
                    //double hprpeng = pc.fracproddict[inst].hprpeng;
                    //if (hprpeng == 0)
                    //    hprpeng = qprog.First().fracproddict[inst].hprpeng;
                    //if (pc.fk)
                    //{
                    double hstperstud = pc.hp / 60;
                    string vtsem = "VT" + year;
                    string htsem = "HT" + year;
                    int colvt = planhd[t1string + vtsem] + 1;
                    int colht = planhd[t1string + htsem] + 1;
                    if (pc.getbatch(vtsem) != null && pc.getbatch(vtsem).getactualstud(1) > 0)
                    {
                        sheet.Cells[row, colvt].Value = pc.getbatch(vtsem).getactualstud(1);
                    }
                    if (pc.getbatch(htsem) != null && pc.getbatch(htsem).getactualstud(1) > 0)
                    {
                        sheet.Cells[row, colht].Value = pc.getbatch(htsem).getactualstud(1);
                    }
                    sheet.Cells[row, colyear].Formula = toreplace + "=" + (hstpeng + pc.totalprod.prestationsgrad() * hprpeng) + "*"+hstperstud + "*(" + Cellname(row, colvt) + "+" + Cellname(row, colht) + ")";
                    //}
                    //else
                    //{
                    //    int colhst = plan2hd[hststring + "20" + year] + 1;
                    //    int colhpr = plan2hd[hprstring + "20" + year] + 1;
                    //    sheet.Cells[row, colyear].Formula = toreplace + "=" + hstpeng + "*'" + detailsheetname + "'!" + Cellname(row, colhst) + "+" + hprpeng + "*'" + detailsheetname + "'!" + Cellname(row, colhpr);
                    //}
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

                sem = startsem;
                do
                {
                    int col = planhd[t1string + sem] + 1;
                    add_batsemref(row, col, pc.name, sem, 1, paketsheetname);
                    sem = util.incrementsemester(sem);
                }
                while (sem != util.incrementsemester(endsem));


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
                if (pc.homeinst == Form1.utaninst)
                    continue;
                int row = rrow[sheet.Name][pc.name];
                bool ht = phtstart[pc.name];
                bool vt = pvtstart[pc.name];
                int icol = lastcolwithdata + 1;
                while (icol <= planhd.Count)
                {
                    if (ht)
                        sheet.Cells[row, icol].Interior.Color = Excel.XlRgbColor.rgbYellow;
                    if (vt && icol < planhd.Count)
                        sheet.Cells[row, icol + 1].Interior.Color = Excel.XlRgbColor.rgbYellow;
                    icol += 2;
                }

            }

            //sheet.FreezeColumns(1);
            //sheet.Protect();
        }

        private void PaketSheetPrognos(Excel.Worksheet sheet, List<programclass> qprog, string startsem, string endsem, string inst)
        {
            SheetWithHeader(sheet, qprog.Count + 2, planhd);

            string prognossem = util.incrementsemester(lastsemwithdata);

            if (!rrow.ContainsKey(sheet.Name))
                rrow.Add(sheet.Name, ProgramNames(sheet, qprog, progoffset));
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
                if (pc.homeinst == Form1.utaninst)
                    continue;
                int row = rrow[sheet.Name][pc.name];
                for (int year = startyear; year <= endyear; year++)
                {
                    int colyear = planhd[moneystring + "20" + year] + 1;
                    string firstcourse = pc.coursedict[1].Keys.First();
                    programclass cc = Form1.fkcodedict[firstcourse];
                    double hstpeng = Form1.hstkr(1, cc.studentpengarea,year);
                    double hprpeng = Form1.hprkr(1, cc.studentpengarea,year);
                    //double hstpeng = pc.fracproddict[inst].hstpeng;
                    //if (hstpeng == 0)
                    //    hstpeng = qprog.First().fracproddict[inst].hstpeng;
                    //double hprpeng = pc.fracproddict[inst].hprpeng;
                    //if (hprpeng == 0)
                    //    hprpeng = qprog.First().fracproddict[inst].hprpeng;
                    //if (pc.fk)
                    //{
                    double hstperstud = pc.hp / 60;
                    int colvt = planhd[t1string + "VT" + year] + 1;
                    int colht = planhd[t1string + "HT" + year] + 1;
                    sheet.Cells[row, colyear].Formula = toreplace + "=" + (hstpeng + pc.totalprod.prestationsgrad() * hprpeng) + "*" + hstperstud + "*(" + Cellname(row, colvt) + "+" + Cellname(row, colht) + ")";
                    //}
                    //else
                    //{
                    //    int colhst = plan2hd[hststring + "20" + year] + 1;
                    //    int colhpr = plan2hd[hprstring + "20" + year] + 1;
                    //    sheet.Cells[row, colyear].Formula = toreplace + "=" + hstpeng + "*'" + detailsheetname + "'!" + Cellname(row, colhst) + "+" + hprpeng + "*'" + detailsheetname + "'!" + Cellname(row, colhpr);
                    //}
                }

                string studhst = pc.fk ? "HST" : "Stud";
                sheet.Cells[row, planhd[studhststring] + 1] = studhst;
                sheet.Cells[row, planhd[inststring] + 1] = Form1.instshortdict[pc.homeinst];

                phtstart.Add(pc.name, false);
                pvtstart.Add(pc.name, false);


                string sem = startsem;
                do
                {
                    int col = planhd[t1string + sem] + 1;
                    add_batsemref(row, col, pc.name, sem, 1, paketsheetname);
                    sem = util.incrementsemester(sem);
                }
                while (sem != util.incrementsemester(endsem));

                //sem = prognossem;
                sem = startsem;
                //if (!pc.fk)
                do
                {
                    int col = planhd["T1 " + sem] + 1;
                    int acccol = planhd[acceptstring + sem] + 1;
                    int u1col = planhd[acceptu1string + sem] + 1;
                    int u2col = planhd[acceptu2string + sem] + 1;
                    int applcol = planhd[applstring + sem] + 1;
                    int retacccol = rethd["Antagen -> T1"] + 1;
                    int retu1col = rethd["U1 -> T1"] + 1;
                    int retu2col = rethd["U2 -> T1"] + 1;
                    int retapplcol = rethd["Sökande -> T1"] + 1;
                    
                    programbatchclass bc = (from c in pc.batchlist where c.batchstart == sem select c).FirstOrDefault();
                    if (bc != null && bc.getstud(1) > 0)
                    {
                        sheet.Cells[row, col] = (double)bc.getstud(1);
                    }
                    else
                    {
                        string f = toreplace + "=IF(" + Cellname(row, acccol) + ">0;" + retpaketsheetname + "!" + Cellname(row, retacccol) + "*" + Cellname(row, acccol) + ";"
                            + "IF(" + Cellname(row, u2col) + " > 0; " + retpaketsheetname + "!" + Cellname(row, retu2col) + "*" + Cellname(row, u2col) + "; "
                            + "IF(" + Cellname(row, u1col) + " > 0; " + retpaketsheetname + "!" + Cellname(row, retu1col) + "*" + Cellname(row, u1col) + "; "
                            + "IF(" + Cellname(row, applcol) + " > 0; " + retpaketsheetname + "!" + Cellname(row, retapplcol) + "*" + Cellname(row, applcol) + ";0))))";
                        sheet.Cells[row, col].Formula = f;
                    }

                    if (bc != null)
                    {
                        if (bc.applicants[0] != null)
                            sheet.Cells[row, applcol] = (double)bc.applicants[0];
                        if (bc.applicants[1] != null)
                            sheet.Cells[row, u1col] = (double)bc.applicants[1];
                        if (bc.applicants[2] != null)
                            sheet.Cells[row, u2col] = (double)bc.applicants[2];
                        if (bc.applicants[3] != null)
                            sheet.Cells[row, acccol] = (double)bc.applicants[3];
                    }

                    if (sem == prognossem)
                    {
                        int budgetcol = planhd[budgetstring + sem] + 1;
                        if (bc != null)
                            sheet.Cells[row, budgetcol] = bc.budget_T1;
                        else
                            sheet.Cells[row, budgetcol] = 0;

                        int diffcol = planhd[diffstring + sem] + 1;
                        sheet.Cells[row, diffcol].Formula = toreplace + "=" + Cellname(row, col) + "-" + Cellname(row, budgetcol);
                    }

                    sem = util.incrementsemester(sem);
                }
                while (sem != util.incrementsemester(endsem));

                sem = startsem;
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
                            int nstud = (int)bc.getstud(1);
                            if (nstud > 0)
                            {
                                sheet.Cells[row, col] = nstud;
                                if (bc.batchstart.Contains("H"))
                                    phtstart[pc.name] = true;
                                else
                                    pvtstart[pc.name] = true;
                            }
                        }
                    }
                    sem = util.incrementsemester(sem);
                }
                while (sem != util.incrementsemester(lastsemwithdata));


            }


            PrognosColors(sheet, qprog, lastcolwithdata, prognossem);
            //sheet.FreezeColumns(1);
            //sheet.Protect();
        }

        string toreplace = "§§§";

        private void DetailSheet(Excel.Worksheet sheet, List<programclass> qprog, string startsem, string endsem,string inst)
        {
            if (!rrow.ContainsKey(sheet.Name))
                rrow.Add(sheet.Name, ProgramNames(sheet, qprog, progoffset));
            sheet.Cells[2, 1] = "Total";
            sheet.Rows[2].Font.Bold = true;
            for (int icol = 5;icol<=plan2hd.Count;icol++)
            {
                sheet.Cells[2,icol].Formula = toreplace+"=SUM("+Cellname(3,icol)+":"+Cellname(3+qprog.Count,icol)+")";
            }

            int meancol = 2;
            int tr0col = retoffset;//3;xxxx
            int prestcol = 4;

            //='Sheet 1'!A3

            foreach (programclass pc in qprog)
            {
                int row = rrow[sheet.Name][pc.name];
                sheet.Cells[row, meancol].Formula = "='" + retsheetname + "'!" + Cellname(row, meancol);
                sheet.Cells[row, tr0col].Formula = "='" + retsheetname + "'!" + Cellname(row, tr0col);
                double prest = pc.prod_per_student.prestationsgrad();
                if (prest > 1)
                    prest = 0.8;
                sheet.Cells[row, prestcol] = prest;

                int batrow = batsheetrow;
                //double frachst = 0.5; // pc.fracproddict[inst].frachst;
                double frachst = pc.fracproddict.ContainsKey(inst) ? pc.fracproddict[inst].frachst : 0;
                //double frachpr = frachst*0.8; // pc.fracproddict[inst].frachst;
                double frachpr = pc.fracproddict.ContainsKey(inst) ? pc.fracproddict[inst].frachpr : 0;
                string semx = startsem;
                do
                {
                    //int retcol = retoffset + nsem;
                    string hsts = "=SUMIF('" + batsheetname + "'!A2:Z" + batrow + ";" + Cellname(row, 1) + ";'" + batsheetname + "'!" + Cellname(2, bathd[semx] + 1) + ":" + Cellname(batrow, bathd[semx] + 1) + ")*" + frachst; //"*'" + retsheetname + "'!" + Cellname(prow[prog], retcol);
                    sheet.Cells[row, plan2hd[hststring + semx] + 1] = toreplace+hsts;
                    //sheet.Cells[row, plan2hd[hststring + semx] + 1] = hsts;
                    string hprs = "=SUMIF('" + batsheetname + "'!A2:Z" + batrow + ";" + Cellname(row, 1) + ";'" + batsheetname + "'!" + Cellname(2, bathd[semx] + 1) + ":" + Cellname(batrow, bathd[semx] + 1) + ")*" + frachpr; //"*'" + retsheetname + "'!" + Cellname(prow[prog], retcol);
                    sheet.Cells[row, plan2hd[hprstring + semx] + 1] = toreplace+hprs;

                    string antags = "='" + mainsheetname + "'!" + Cellname(row, planhd[t1string + semx] + 1)+"/'"+retsheetname+"'!"+Cellname(row,tr0col);
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
            //sheet.Cells[1, 1].Locked = false;
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

            foreach (string s in plan2hd.Keys)
            {
                int ncol = plan2hd[s] + 1;
                Color color = Color.White;
                if (s.Contains(acceptstring))
                    color = Color.LightGreen;
                else if (s.Contains(hststring))
                    color = Color.LightBlue;
                else if (s.Contains(hprstring))
                    color = Color.PaleTurquoise;
                sheet.Range[Cellname(3, ncol), Cellname(qprog.Count()+2, ncol)].Interior.Color = color;
            }

            //sheet.FreezeColumns(1);
            //sheet.Protect();
        }

        private void TriangelSheet(Excel.Worksheet triangelsheet, string startsem, string endsem)
        {
            triangelhd = new Dictionary<string, int>() { { "Triangel/ämne", 0 } };

            int startyear = 2000 + util.semtoint(startsem);
            int endyear = 2000 + util.semtoint(endsem);
            int col = triangelhd.Count;

            for (int i = startyear; i <= endyear; i++)
            {
                triangelhd.Add(moneystring + i, col);
                col++;
            }

            for (int i = startyear; i <= endyear; i++)
            {
                triangelhd.Add(moneystring + fkstring + i, col);
                col++;
            }

            for (int i = startyear; i <= endyear; i++)
            {
                triangelhd.Add(moneystring + progstring + i, col);
                col++;
            }

            int totalrow = 4;
            int takbelopprow = 2;
            int diffrow = 3;

            List<string> triangles = Form1.trianglecolor.Keys.ToList();

            SheetWithHeader(triangelsheet, triangles.Count + Form1.subjcodetriangle.Count + triangeloffset, triangelhd);

            int subjstartrow = totalrow + triangles.Count + 4;
            int subjtotalrow = subjstartrow - 1;
            int srow = subjstartrow;
            Dictionary<string, int> subjrowdict = new Dictionary<string, int>();

            for (int i=2;i<=triangelhd.Count;i++)
            {
                triangelsheet.Cells[subjstartrow - 2, i] = triangelsheet.Cells[1, i];
            }
            foreach (string subj in Form1.subjcodetriangle.Keys)
            {
                triangelsheet.Cells[srow, 1] = subj;
                subjrowdict.Add(subj, srow);

                for (int i = startyear; i <= endyear; i++)
                {
                    string kr = moneystring + i;
                    string krfk = moneystring + fkstring + i;
                    string krprog = moneystring + progstring + i;

                    int fkrow = fksumrow[subj];

                    triangelsheet.Cells[srow, triangelhd[kr] + 1] = toreplace + "=" +
                        Cellname(srow, triangelhd[krfk] + 1) + "+" +
                        Cellname(srow, triangelhd[krprog] + 1);

                    triangelsheet.Cells[srow, triangelhd[krfk] + 1] = toreplace + "=0,001*'" +
                        coursesheetname + "'!" + Cellname(fksumrow[subj], coursehd[krfk] + 1);
                    triangelsheet.Cells[srow, triangelhd[krprog] + 1] = toreplace + "=0,001*'" +
                        coursesheetname + "'!" + Cellname(fksumrow[subj], coursehd[krprog] + 1);
                }

                srow++;
            }

            int trow = totalrow + 1;
            foreach (string tri in triangles)
            {
                triangelsheet.Cells[trow, 1] = tri;
                List<string> subjlist = (from c in Form1.subjcodetriangle where c.Value == tri select c.Key).ToList();
                for (int i = 2; i <= triangelhd.Count; i++)
                {
                    StringBuilder sb = new StringBuilder(toreplace + "=");
                    string plus = "";
                    foreach (string subj in subjlist)
                    {
                        sb.Append(plus + Cellname(subjrowdict[subj], i));
                        plus = "+";
                    }
                    triangelsheet.Cells[trow, i] = sb.ToString();
                }

                trow++;
            }

            triangelsheet.Cells[totalrow, 1] = "Total";
            triangelsheet.Rows[totalrow].Font.Bold = true;
            triangelsheet.Cells[takbelopprow, 1] = "Takbelopp";
            triangelsheet.Cells[diffrow, 1] = "Över/underproduktion";
            triangelsheet.Cells[subjtotalrow, 1] = "Total";
            triangelsheet.Rows[subjtotalrow].Font.Bold = true;

            for (int i=startyear;i<=endyear;i++)
            {
                if (!Form1.takbelopp.ContainsKey(i))
                    continue;
                string kr = moneystring + i;
                int col1 = triangelhd[kr] + 1;
                triangelsheet.Cells[takbelopprow, col1].Value = Form1.takbelopp[i];
                triangelsheet.Cells[diffrow, col1] = toreplace + "=" + Cellname(totalrow, col1) + "-" + Cellname(takbelopprow, col1);
            }

            for (int i = 2; i <= triangelhd.Count; i++)
            {
                triangelsheet.Cells[totalrow, i] = toreplace + "=SUM(" + Cellname(totalrow + 1, i) + ":" + Cellname(trow - 1, i) + ")";
                triangelsheet.Cells[subjtotalrow, i] = toreplace + "=SUM(" + Cellname(subjtotalrow + 1, i) + ":" + Cellname(srow - 1, i) + ")";

                Excel.Range qa = triangelsheet.Columns[i];
                qa.ColumnWidth = 11;
                qa.NumberFormat = "# ###";
                if (triangelsheet.Cells[1,i].Value.ToString().Contains("prog"))
                    qa.Interior.Color = Excel.XlRgbColor.rgbPink;
                else if (triangelsheet.Cells[1, i].Value.ToString().Contains("FK"))
                    qa.Interior.Color = Excel.XlRgbColor.rgbLightPink;
                else
                    qa.Interior.Color = Excel.XlRgbColor.rgbMistyRose;
            }
            Excel.Range qb = triangelsheet.Rows[trow];
            qb.Interior.Color = Excel.XlRgbColor.rgbWhite;


        }

        private void SumSheet(Excel.Worksheet sumsheet, List<programclass> qprog, string startsem, string endsem)
        {
            sumhd = new Dictionary<string, int>() { { "Prog/inst", 0 } };

            int startyear = 2000 + util.semtoint(startsem);
            int endyear = 2000 + util.semtoint(endsem);
            int col = sumhd.Count;

            for (int i = startyear; i <= endyear; i++)
            {
                sumhd.Add(moneystring + i, col);
                col++;
            }

            for (int i = startyear; i <= endyear; i++)
            {
                sumhd.Add(instsumstring + i, col);
                col++;
            }

            foreach (string inst in Form1.institutiondict.Keys)
            {
                for (int i = startyear; i <= endyear; i++)
                {
                    sumhd.Add(Form1.instshortdict[inst] + " " + i, col);
                    col++;
                }
            }

            int totalrow = 2;

            SheetWithHeader(sumsheet, qprog.Count + sumoffset, sumhd);
            if (!rrow.ContainsKey(sumsheet.Name))
                rrow.Add(sumsheet.Name, ProgramNames(sumsheet, qprog, sumoffset));

            sumsheet.Cells[totalrow, 1] = "Total";
            sumsheet.Rows[totalrow].Font.Bold = true;

            //"='" + mainsheetname + "'!" + Cellname(row, planhd[t1string + semx] + 1)+"/'"+retsheetname+"'!"+Cellname(row,tr0col);

            for (int i = startyear; i <= endyear; i++)
            {
                int coltot = sumhd[moneystring + i] + 1;
                int colinstsum = sumhd[instsumstring + i] + 1;
                string instsum = toreplace + "=";

                foreach (string inst in Form1.institutiondict.Keys)
                {
                    int colinst = sumhd[Form1.instshortdict[inst] + " " + i]+1;
                    instsum += Cellname(totalrow, colinst) + "+";
                    string instsheetname = Form1.instshortdict[inst] + " " + mainsheetname;
                    sumsheet.Cells[totalrow, colinst] = toreplace + "='" + instsheetname + "'!"
                        + Cellname(totalrow, planhd[moneystring + i] + 1);
                }

                sumsheet.Cells[totalrow, coltot] = toreplace + "='" + mainsheetname + "'!"
                    + Cellname(totalrow, planhd[moneystring + i] + 1);

                instsum += "0";
                sumsheet.Cells[totalrow, colinstsum] = instsum;
            }

            foreach (programclass pc in qprog)
            {
                int sumrow = rrow[sumsheet.Name][pc.name];
                int planrow = rrow[mainsheetname][pc.name];

                for (int i = startyear; i <= endyear; i++)
                {
                    int coltot = sumhd[moneystring + i] + 1;
                    int colinstsum = sumhd[instsumstring + i] + 1;
                    string instsum = toreplace + "=";

                    foreach (string inst in Form1.institutiondict.Keys)
                    {
                        int colinst = sumhd[Form1.instshortdict[inst] + " " + i] +1;
                        instsum += Cellname(sumrow, colinst) + "+";
                        string instsheetname = Form1.instshortdict[inst] + " " + mainsheetname;
                        if (rrow[instsheetname].ContainsKey(pc.name))
                        {
                            int instrow = rrow[instsheetname][pc.name];
                            sumsheet.Cells[sumrow, colinst] = toreplace + "='" + instsheetname + "'!"
                                + Cellname(instrow, planhd[moneystring + i] + 1);
                        }
                    }

                    sumsheet.Cells[sumrow, coltot] = toreplace + "='" + mainsheetname + "'!"
                        + Cellname(planrow, planhd[moneystring + i] + 1);

                    instsum += "0";
                    sumsheet.Cells[sumrow, colinstsum] = instsum;
                }


            }

            foreach (string s in sumhd.Keys)
            {
                if (s.Contains("20")) 
                {
                    Excel.Range qa = sumsheet.Columns[sumhd[s] + 1];
                    qa.ColumnWidth = 11;
                    qa.NumberFormat = "# ### ###";
                }

                if (s.Contains(startyear.ToString()))
                {
                    Excel.Range qa = sumsheet.Columns[sumhd[s] + 1];
                    qa.Interior.Color = Color.LightPink;

                }
                else if (s.Contains(endyear.ToString()))
                {
                    Excel.Range qa = sumsheet.Columns[sumhd[s] + 1];
                    qa.Interior.Color = Color.Pink;

                }


            }

        }

        private void fill_coursehd(Excel.Worksheet coursesheet, List<programclass> qprog, List<programclass> qpaket, List<programclass> qcourse, string startsem, string endsem)
        {
            int hpcol = 2;
            int codecol = 3;
            int prestcol = 5;
            int moneycol = 6;

            int totalrow = 2;

            int startyear = 2000 + util.semtoint(startsem);
            int endyear = 2000 + util.semtoint(endsem);

            coursehd = new Dictionary<string, int>() {
                { "Kurs", 0 },
                { "Hp", hpcol-1 },
                { "Kurskod", codecol-1 },
                { "Ämneskod", codecol },
                { prestationstring, prestcol-1 },
                { "Kr/HST", moneycol-1 } };

            List<string> semlist = new List<string>();
            int col = coursehd.Count;


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

            for (int i = startyear; i <= endyear; i++)
            {
                foreach (string uo in Form1.utbomrlist)
                {
                    coursehd.Add(hststring + fkstring + uo + " " + i, col);
                    col++;
                }
            }


        }

        private void CourseSheet(Excel.Worksheet coursesheet, List<programclass> qprog, List<programclass> qpaket, List<programclass> qcourse, string startsem, string endsem)
        {
            int hpcol = 2;
            int codecol = 3;
            int prestcol = 5;
            int moneycol = 6;

            int totalrow = 2;

            int startyear = 2000 + util.semtoint(startsem);
            int endyear = 2000 + util.semtoint(endsem);

            int col = coursehd.Count;

            fksumrow = new Dictionary<string, int>();

            string sem = startsem;
            List<string> semlist = new List<string>();

            do
            {
                semlist.Add(sem);
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));


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
            foreach (programclass pc in qpaket)
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

            SheetWithHeader(coursesheet, qcourse.Count + courseoffset, coursehd);
            if (!rrow.ContainsKey(coursesheet.Name))
                rrow.Add(coursesheet.Name, ProgramNames(coursesheet, qcourse, courseoffset));
            coursesheet.Cells[totalrow, 1] = "Total";
            coursesheet.Rows[totalrow].Font.Bold = true;

            double prestsum = 0;
            double krhstsum = 0;
            double ncourses = 0;

            int firstnewcourseline = -1;
            int lastnewcourseline = -1;

            foreach (programclass pc in qcourse)
            {
                bool nykurs = pc.name.StartsWith(newcoursename);
                int nrow = rrow[coursesheet.Name][pc.name];
                if (!nykurs)
                {
                    coursesheet.Cells[nrow, hpcol].Value = pc.hp < 0 ? 7.5 : pc.hp;
                    if (pc.hp <= 0)
                    {
                        coursesheet.Cells[nrow, hpcol].Borders.Color = Color.Red;
                        string sp = "";
                        if (pc.studentpengarea.Count > 0)
                        {
                            sp = pc.studentpengarea.First().Key + " " + (100 * pc.studentpengarea.First().Value).ToString("N0");
                        }
                        memo(pc.bestcode() + "\t" + pc.hp + "\t" + pc.name + "\t"+sp);
                    }
                }
                else
                {
                    if (firstnewcourseline < 0)
                        firstnewcourseline = nrow;
                    lastnewcourseline = nrow;
                }
                coursesheet.Cells[nrow, codecol].Value = pc.bestcode();
                coursesheet.Cells[nrow, codecol+1].Value = pc.subjectcode;
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
                    krhst = Form1.hstkr(1, pc.studentpengarea,-1) + prest * Form1.hprkr(1, pc.studentpengarea,-1);
                coursesheet.Cells[nrow, moneycol].Value = krhst;
                krhstsum += krhst;
                ncourses++;

                //FK-studenter:
                double lastvt = 0;
                double lastht = 0;
                bool torepeat = false;
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
                        if (sm == lastsemwithdata || sm == util.decrementsemester(lastsemwithdata) || sm == util.incrementsemester(lastsemwithdata))
                            torepeat = true;
                    }
                    else if (torepeat)
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
                        f += (Form1.hstkr(1, pc.studentpengarea,i)) + "*(" + Cellname(nrow, hstvtcol) + "+" + Cellname(nrow, hsthtcol) + ")+";
                        f += (Form1.hprkr(1, pc.studentpengarea,i)) + "*(" + Cellname(nrow, hprvtcol) + "+" + Cellname(nrow, hprhtcol) + ")";
                    }
                    coursesheet.Cells[nrow, ncol].Formula = f;

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
                    f += (Form1.hstkr(1, pc.studentpengarea,i)) + "*(" + Cellname(nrow, hstvtcol) + "+" + Cellname(nrow, hsthtcol) + ")+";
                    f += (Form1.hprkr(1, pc.studentpengarea,i)) + "*(" + Cellname(nrow, hprvtcol) + "+" + Cellname(nrow, hprhtcol) + ")";
                    coursesheet.Cells[nrow, ncol].Formula = f;

                }

                //Summa HST/HPR per år
                for (int i = startyear; i <= endyear; i++)
                {
                    int ncolhst = coursehd[hststring + i] + 1;
                    int ncolhpr = coursehd[hprstring + i] + 1;
                    string vtsem = "VT" + (i % 100);
                    string htsem = "HT" + (i % 100);
                    int hstvtcol = coursehd[hststring +vtsem] + 1;
                    int hsthtcol = coursehd[hststring +htsem] + 1;
                    int hprvtcol = coursehd[hprstring +vtsem] + 1;
                    int hprhtcol = coursehd[hprstring +htsem] + 1;
                    string f = toreplace + "=";

                    coursesheet.Cells[nrow, ncolhst].Formula = f + Cellname(nrow,hstvtcol)+"+"+Cellname(nrow,hsthtcol);
                    coursesheet.Cells[nrow, ncolhpr].Formula = f + Cellname(nrow, hprvtcol) + "+" + Cellname(nrow, hprhtcol);

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
                if (nrow % 100 == 0)
                    memo(nrow.ToString());


                for (int i = startyear; i <= endyear; i++)
                {
                    foreach (string uo in pc.studentpengarea.Keys)
                    {
                        string hd = hststring + fkstring + uo + " " + i;
                        int ncol = coursehd[hd] + 1;
                        string vtsem = "VT" + (i % 100);
                        string htsem = "HT" + (i % 100);
                        int hstvtcol = coursehd[hststring + fkstring + vtsem] + 1;
                        int hsthtcol = coursehd[hststring + fkstring + htsem] + 1;
                        coursesheet.Cells[nrow, ncol].Formula = toreplace + "=" + pc.studentpengarea[uo] + "*(" + Cellname(nrow, hstvtcol) + "+" + Cellname(nrow, hsthtcol) + ")";
                    }
                }

            }
            //sdfa

            coursesheet.Cells[totalrow, 2].Value = "Rödgräns:";
            coursesheet.Cells[totalrow, 3].Value = 10;
            coursesheet.Cells[totalrow, 3].Locked = false;
            for (int icol=7;icol<=coursehd.Count;icol++)
            {
                coursesheet.Cells[totalrow,icol].Formula = toreplace+"=SUM("+Cellname(totalrow+1,icol)+":"+Cellname(qcourse.Count+courseoffset,icol)+")";
            }


            ////Testing if it works to just move it here from CoursesheetPrognos

            var qsubj = (from c in qcourse select c.subjectcode).Distinct();
            var qinst = (from c in qcourse select c.homeinst).Distinct();

            int srow = lastnewcourseline + 2;
            int subjcol = moneycol - 3;
            coursesheet.Cells[srow, subjcol] = "Summerat per ämne";
            //coursesheet.Cells[srow, subjcol + 2] = "Prog+" + coursesheet.Cells[1, moneycol + 3].Value;
            for (int i = subjcol + 4; i <= coursehd.Count; i++)
                coursesheet.Cells[srow, i] = coursesheet.Cells[1, i];
            string coderange = Cellname(totalrow + 1, codecol + 1) + ":" + Cellname(lastnewcourseline, codecol + 1);
            foreach (string subj in qsubj)
            {
                srow++;
                fksumrow.Add(subj, srow);
                coursesheet.Cells[srow, subjcol+1] = subj;
                for (int i = subjcol+4; i <= coursehd.Count; i++)
                    coursesheet.Cells[srow, i] = toreplace + "=SUMIF(" + coderange + ";" + Cellname(srow, subjcol+1) + ";" + Cellname(totalrow + 1, i) + ":" + Cellname(lastnewcourseline, i) + ")";
                //coursesheet.Cells[srow,subjcol+2] = toreplace + "=" + Cellname(srow, moneycol + 3) + "+" + Cellname(srow, moneycol+8);
                //coursesheet.Cells[srow, subjcol+3] = toreplace + "=" + Cellname(srow, instcol + 4) + "+" + Cellname(srow, instcol + 6);
            }

            //srow++;
            //coursesheet.Cells[srow, instcol] = "Summerat per institution";
            //string instrange = Cellname(totalrow + 1, instcol) + ":" + Cellname(lastnewcourseline, instcol);
            //foreach (string inst in qinst)
            //{
            //    srow++;
            //    fksumrow.Add(inst, srow);
            //    coursesheet.Cells[srow, instcol] = inst;
            //    for (int i = moneycol + 1; i < coursehd.Count; i++)
            //        coursesheet.Cells[srow, i] = toreplace + "=SUMIF(" + instrange + ";" + Cellname(srow, instcol) + ";" + Cellname(totalrow + 1, i) + ":" + Cellname(lastnewcourseline, i) + ")";
            //    coursesheet.Cells[srow, instcol + 1] = toreplace + "=" + Cellname(srow, instcol + 3) + "+" + Cellname(srow, instcol + 5);
            //    coursesheet.Cells[srow, instcol + 2] = toreplace + "=" + Cellname(srow, instcol + 4) + "+" + Cellname(srow, instcol + 6);
            //}

            ////End testing section


            coursesheet.Range["e3", Cellname(qcourse.Count + courseoffset, 5)].NumberFormat = "###.0%";
            coursesheet.Range["f3", "p"+srow].NumberFormat = "# ### ###";
            coursesheet.Range["f3", "p"+lastnewcourseline].ColumnWidth = 13;
            coursesheet.Range["e"+(lastnewcourseline+2), "e" + srow].NumberFormat = "# ### ###";
            coursesheet.Range["A1", "GD1"].WrapText = true;
            //coursesheet.Range["f3", "O999"].Interior.Color = Color.Pink;
            ////coursesheet.Range["z3", "ai999"].NumberFormat = "# ###.#";
            //coursesheet.Range["z3", "ai999"].Interior.Color = Color.Yellow;
            coursesheet.Range["aj3", "dt"+lastnewcourseline].NumberFormat = "# ###.#";
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
                {
                    qa.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    qa.Locked = false;
                }
                if(color == Color.AntiqueWhite)
                {
                    Excel.FormatCondition cond = qa.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlBetween,0.01, "=$c$2");
                    cond.Font.Color = Color.Red;
                    qa.NumberFormat = "###";
                }
            }

            if (firstnewcourseline > 0 && lastnewcourseline >= firstnewcourseline)
            {
                coursesheet.Range[Cellname(firstnewcourseline, 1), Cellname(lastnewcourseline, 3)].Interior.Color = Color.Yellow;
                coursesheet.Range[Cellname(firstnewcourseline, 1), Cellname(lastnewcourseline, 3)].Locked = false;
            }

        }

        private void CourseSheetPrognos(Excel.Worksheet coursesheet, List<programclass> qprog, List<programclass> qpaket, List<programclass> qcourse, string startsem, string endsem)
        {
            int hpcol = 2;
            int codecol = 3;
            int instcol = 5;
            int prestcol = 6;
            int moneycol = 7;
            int totalrow = 2;

            fksumrow = new Dictionary<string, int>();

            string prognossem = util.incrementsemester(lastsemwithdata);

            coursehd = new Dictionary<string, int>() {
                { "Kurs", 0 },
                { "Hp", hpcol-1 },
                { "Kurskod", codecol-1 },
                { "Ämneskod", codecol },
                { "Inst", instcol-1 },
                { prestationstring, prestcol-1 },
                { "Kr/HST", moneycol-1 } };

            List<string> semlist = new List<string>();
            int col = coursehd.Count;

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
            //col = coursehd.Count + 1;
            do
            {
                coursehd.Add(applstring + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));

            sem = startsem;
            //col = coursehd.Count + 1;
            do
            {
                coursehd.Add(acceptu1string + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));

            sem = startsem;
            //col = coursehd.Count + 1;
            do
            {
                coursehd.Add(acceptu2string + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));

            sem = startsem;
            //col = coursehd.Count + 1;
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

            sem = endsem;
            coursehd.Add(budgetstring + sem, col);
            col++;
            coursehd.Add(diffstring + sem, col);
            col++;

            sem = startsem;

            do
            {
                coursehd.Add(progstudstring + sem, col);
                col++;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(endsem));

            //sem = startsem;

            //do
            //{
            //    coursehd.Add(acceptstring + sem, col);
            //    col++;
            //    sem = util.incrementsemester(sem);
            //}
            //while (sem != util.incrementsemester(endsem));


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
            foreach (programclass pc in qpaket)
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

            SheetWithHeader(coursesheet, qcourse.Count + courseoffset, coursehd);
            if (!rrow.ContainsKey(coursesheet.Name))
                rrow.Add(coursesheet.Name, ProgramNames(coursesheet, qcourse, courseoffset));
            coursesheet.Cells[totalrow, 1] = "Total";
            coursesheet.Rows[totalrow].Font.Bold = true;
            coursesheet.Cells[totalrow, 1].AddComment("Test comment");

            double prestsum = 0;
            double krhstsum = 0;
            double ncourses = 0;

            int firstnewcourseline = -1;
            int lastnewcourseline = -1;

            foreach (programclass pc in qcourse)
            {
                bool nykurs = pc.name.StartsWith(newcoursename);
                int nrow = rrow[coursesheet.Name][pc.name];
                if (!nykurs)
                {
                    coursesheet.Cells[nrow, hpcol].Value = pc.hp;
                    if (pc.hp <= 0)
                    {
                        string sp = "";
                        if (pc.studentpengarea.Count > 0)
                        {
                            sp = pc.studentpengarea.First().Key + " " + (100 * pc.studentpengarea.First().Value).ToString("N0");
                        }
                        memo(pc.bestcode() + "\t" + pc.hp + "\t" + pc.name + "\t" + sp);
                    }
                }
                else
                {
                    if (firstnewcourseline < 0)
                        firstnewcourseline = nrow;
                    lastnewcourseline = nrow;
                }
                coursesheet.Cells[nrow, codecol].Value = pc.bestcode();
                coursesheet.Cells[nrow, codecol + 1].Value = pc.subjectcode;
                coursesheet.Cells[nrow, instcol].Value = Form1.instshortdict[pc.homeinst];
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
                    krhst = Form1.hstkr(1, pc.studentpengarea,-1) + prest * Form1.hprkr(1, pc.studentpengarea,-1);
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
                        double? stud = bc.getactualstud(1);
                        if (stud != null && stud > 0)
                        {
                            coursesheet.Cells[nrow, ncol].Value = (double)stud;
                            if (sm.StartsWith("V"))
                                lastvt = (double)stud;
                            else
                                lastht = (double)stud;
                        }
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
                    int ncol = coursehd[hststring + fkstring + sm] + 1;
                    int ncolf = coursehd[fkstudstring + sm] + 1;

                    coursesheet.Cells[nrow, ncol].Formula = toreplace + "=" + Cellname(nrow, ncolf)
                        + "*" + Cellname(nrow, hpcol) + "/60";
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
                    int ncolf = coursehd[hststring + fkstring + sm] + 1;
                    int ncolp = coursehd[hststring + progstring + sm] + 1;
                    coursesheet.Cells[nrow, ncol].Formula = toreplace + "=" + Cellname(nrow, ncolf) + "+" + Cellname(nrow, ncolp);
                }


                //HPR FK
                foreach (string sm in semlist)
                {
                    int ncol = coursehd[hprstring + fkstring + sm] + 1;
                    int ncolf = coursehd[hststring + fkstring + sm] + 1;

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
                    int ncol = coursehd[moneystring + fkstring + i] + 1;
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
                        f += Form1.hstkr(1, pc.studentpengarea,i) + "*(" + Cellname(nrow, hstvtcol) + "+" + Cellname(nrow, hsthtcol) + ")+";
                        f += Form1.hprkr(1, pc.studentpengarea,i) + "*(" + Cellname(nrow, hprvtcol) + "+" + Cellname(nrow, hprhtcol) + ")";
                    }
                    coursesheet.Cells[nrow, ncol].Formula = f;

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
                    f += Form1.hstkr(1, pc.studentpengarea,i) + "*(" + Cellname(nrow, hstvtcol) + "+" + Cellname(nrow, hsthtcol) + ")+";
                    f += Form1.hprkr(1, pc.studentpengarea,i) + "*(" + Cellname(nrow, hprvtcol) + "+" + Cellname(nrow, hprhtcol) + ")";
                    coursesheet.Cells[nrow, ncol].Formula = f;

                }

                //Summa HST/HPR per år
                for (int i = startyear; i <= endyear; i++)
                {
                    int ncolhst = coursehd[hststring + i] + 1;
                    int ncolhpr = coursehd[hprstring + i] + 1;
                    string vtsem = "VT" + (i % 100);
                    string htsem = "HT" + (i % 100);
                    int hstvtcol = coursehd[hststring + vtsem] + 1;
                    int hsthtcol = coursehd[hststring + htsem] + 1;
                    int hprvtcol = coursehd[hprstring + vtsem] + 1;
                    int hprhtcol = coursehd[hprstring + htsem] + 1;
                    string f = toreplace + "=";

                    coursesheet.Cells[nrow, ncolhst].Formula = f + Cellname(nrow, hstvtcol) + "+" + Cellname(nrow, hsthtcol);
                    coursesheet.Cells[nrow, ncolhpr].Formula = f + Cellname(nrow, hprvtcol) + "+" + Cellname(nrow, hprhtcol);

                }

                //Antas
                //foreach (string sm in semlist)
                //{
                //    int ncol = coursehd[acceptstring + sm] + 1;
                //    int ncolf = coursehd[fkstudstring + sm] + 1;
                //    double transition = 0.666;
                //    if (pc.transition[0] != null && pc.transition[0].transitionprob > 0)
                //        transition = pc.transition[0].transitionprob;
                //    coursesheet.Cells[nrow, ncol].Formula = toreplace + "=IF(" + Cellname(nrow, ncolf) + ">0;" + Cellname(nrow, ncolf) + "/" + transition + ";\"\")";
                //}

                sem = prognossem;
                //if (!pc.fk)
                {
                    int plancol = coursehd[fkstudstring + sem] + 1;
                    int acccol = coursehd[acceptstring + sem] + 1;
                    int u1col = coursehd[acceptu1string + sem] + 1;
                    int u2col = coursehd[acceptu2string + sem] + 1;
                    int applcol = coursehd[applstring + sem] + 1;
                    int retacccol = rethd["Antagen -> T1"] + 1;
                    int retu1col = rethd["U1 -> T1"] + 1;
                    int retu2col = rethd["U2 -> T1"] + 1;
                    int retapplcol = rethd["Sökande -> T1"] + 1;

                    programbatchclass bc = (from c in pc.batchlist where c.batchstart == sem select c).FirstOrDefault();
                    if (bc != null && bc.getstud(1) > 0)
                    {
                        coursesheet.Cells[nrow, plancol] = (double)bc.getstud(1);
                    }
                    else
                    {
                        string f = toreplace + "=IF(" + Cellname(nrow, acccol) + ">0;" + retcoursesheetname + "!" + Cellname(nrow, retacccol) + "*" + Cellname(nrow, acccol) + ";"
                            + "IF(" + Cellname(nrow, u2col) + " > 0; " + retcoursesheetname + "!" + Cellname(nrow, retu2col) + "*" + Cellname(nrow, u2col) + "; "
                            + "IF(" + Cellname(nrow, u1col) + " > 0; " + retcoursesheetname + "!" + Cellname(nrow, retu1col) + "*" + Cellname(nrow, u1col) + "; "
                            + "IF(" + Cellname(nrow, applcol) + " > 0; " + retcoursesheetname + "!" + Cellname(nrow, retapplcol) + "*" + Cellname(nrow, applcol) + ";0))))";
                        coursesheet.Cells[nrow, plancol].Formula = f;
                    }

                    if (bc != null)
                    {
                        if (bc.applicants[0] != null)
                            coursesheet.Cells[nrow, applcol] = (double)bc.applicants[0];
                        if (bc.applicants[1] != null)
                            coursesheet.Cells[nrow, u1col] = (double)bc.applicants[1];
                        if (bc.applicants[2] != null)
                            coursesheet.Cells[nrow, u2col] = (double)bc.applicants[2];
                        if (bc.applicants[3] != null)
                            coursesheet.Cells[nrow, acccol] = (double)bc.applicants[3];
                    }

                    int budgetcol = coursehd[budgetstring + sem] + 1;
                    if (bc != null)
                        coursesheet.Cells[nrow, budgetcol] = bc.budget_T1;
                    else
                        coursesheet.Cells[nrow, budgetcol] = 0;

                    int diffcol = coursehd[diffstring + sem] + 1;
                    coursesheet.Cells[nrow, diffcol].Formula = toreplace + "=" + Cellname(nrow, plancol) + "-" + Cellname(nrow, budgetcol);

                }
                if (nrow % 100 == 0)
                    memo(nrow.ToString());

            }
            //sdfa

            coursesheet.Cells[totalrow, 2].Value = "Rödgräns:";
            coursesheet.Cells[totalrow, 3].Value = 10;
            coursesheet.Cells[totalrow, 3].Locked = false;
            for (int icol = 7; icol <= coursehd.Count; icol++)
            {
                coursesheet.Cells[totalrow, icol].Formula = toreplace + "=SUM(" + Cellname(totalrow + 1, icol) + ":" + Cellname(qcourse.Count + courseoffset, icol) + ")";
            }

            var qsubj = (from c in qcourse select c.subjectcode).Distinct();
            var qinst = (from c in qcourse select c.homeinst).Distinct();

            int srow = lastnewcourseline + 2;
            coursesheet.Cells[srow, 1] = "Summerat per ämne";
            string coderange = Cellname(totalrow + 1, codecol + 1)+":"+Cellname(lastnewcourseline,codecol+1);
            foreach (string subj in qsubj)
            {
                srow++;
                fksumrow.Add(subj, srow);
                coursesheet.Cells[srow, instcol] = subj;
                for (int i = moneycol + 1; i <= coursehd.Count; i++)
                    coursesheet.Cells[srow, i] = toreplace + "=SUMIF(" + coderange + ";" + Cellname(srow, instcol) + ";"+ Cellname(totalrow + 1, i) + ":" + Cellname(lastnewcourseline, i) + ")";
                coursesheet.Cells[srow, instcol + 1] = toreplace + "=" + Cellname(srow, instcol + 3) + "+" + Cellname(srow, instcol + 5);
                coursesheet.Cells[srow, instcol + 2] = toreplace + "=" + Cellname(srow, instcol + 4) + "+" + Cellname(srow, instcol + 6);
            }

            srow++;
            coursesheet.Cells[srow, instcol] = "Summerat per institution";
            string instrange = Cellname(totalrow + 1, instcol) + ":" + Cellname(lastnewcourseline, instcol);
            foreach (string inst in qinst)
            {
                srow++;
                fksumrow.Add(inst, srow);
                coursesheet.Cells[srow, instcol] = Form1.instshortdict[inst];
                for (int i = moneycol + 1; i < coursehd.Count; i++)
                    coursesheet.Cells[srow, i] = toreplace + "=SUMIF(" + instrange + ";" + Cellname(srow, instcol) + ";" + Cellname(totalrow + 1, i) + ":" + Cellname(lastnewcourseline, i) + ")";
                coursesheet.Cells[srow, instcol + 1] = toreplace + "=" + Cellname(srow, instcol + 3) + "+" + Cellname(srow, instcol + 5);
                coursesheet.Cells[srow, instcol + 2] = toreplace + "=" + Cellname(srow, instcol + 4) + "+" + Cellname(srow, instcol + 6);
            }




            coursesheet.Range["D3", Cellname(srow, 4)].NumberFormat = "###.0%";
            coursesheet.Range["g3", "p"+srow].NumberFormat = "# ### ###";
            coursesheet.Range["g3", "p"+srow].ColumnWidth = 13;
            //coursesheet.Range["f3", "O999"].Interior.Color = Color.Pink;
            ////coursesheet.Range["z3", "ai999"].NumberFormat = "# ###.#";
            //coursesheet.Range["z3", "ai999"].Interior.Color = Color.Yellow;
            coursesheet.Range["aj3", "dt"+srow].NumberFormat = "# ###.#";
            //coursesheet.Range["aj3", "as999"].Interior.Color = Color.Tan;

            coursesheet.Range[Cellname(lastnewcourseline, instcol), Cellname(srow, instcol)].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

            int lastcolwithdata = -1;
            sem = startsem;
            do
            {
                lastcolwithdata = coursehd[fkstudstring + sem] + 1;
                sem = util.incrementsemester(sem);
            }
            while (sem != util.incrementsemester(lastsemwithdata));

            PrognosColorsCourse(coursesheet, qcourse, lastcolwithdata, prognossem);

            coursesheet.Range[Cellname(firstnewcourseline, 1), Cellname(lastnewcourseline, 3)].Interior.Color = Color.Yellow;
            coursesheet.Range[Cellname(firstnewcourseline, 1), Cellname(lastnewcourseline, 3)].Locked = false;

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
                        if (bc.applicants[0] == null)
                            batappl.Add(bc.batchstart, 0);
                        else
                            batappl.Add(bc.batchstart, (double)bc.applicants[0]);
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
            //if (!rrow.ContainsKey(sheet.Name))
            //    rrow.Add(sheet.Name, ProgramNames(sheet, qprog, progoffset));

            string prognossem = util.incrementsemester(lastsemwithdata);

            nrow = 1;
            foreach (string prog in progbatsem.Keys)
            {
                int nsemtot = Form1.origprogramdict[prog].semesters;

                foreach (string bat in progbatsem[prog].Keys)
                {
                    //if (prognos && bat == prognossem) //kluge to get around duplicate lines 240902
                    //    continue;
                    if (bat == prognossem) //kluge to get around duplicate lines 240904
                        continue;

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
                            sheet.Cells[nrow, bathd[semx] + 1].Formula = "=" + Cellname(nrow, bathd[semx])+ "*'" + retsheetname + "'!" + Cellname(rrow[retsheetname][prog], retcol);
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
                    sheet.Cells[nrow, bathd[semnewbatch] + 1].Formula = "='" + mainsheetname + "'!" + Cellname(rrow[mainsheetname][prog], planhd[t1string + semnewbatch] + 1);
                    add_batsemref(nrow, bathd[semnewbatch] + 1, prog, semnewbatch, 1, batsheetname);
                    int nsem = 2;
                    string semx = util.incrementsemester(semnewbatch);
                    if (semnewbatch != endsem)
                    {
                        do
                        {
                            int retcol = retoffset + nsem;
                            sheet.Cells[nrow, bathd[semx] + 1].Formula = "=" + Cellname(nrow, bathd[semx]) + "*'" + retsheetname + "'!" + Cellname(rrow[retsheetname][prog], retcol);
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

                batsheetrow = nrow;
            }
            Excel.Range qa = sheet.Columns[1];
            qa.ColumnWidth = 50;
            //sheet.Protect();
        }

        int batsheetrow = 0;

        private void printfracprod(List<programclass> qprog, string inst)
        {
            memo("\t" + fracprodclass.printheader());
            foreach (programclass pc in qprog)
            {
                memo(pc.name + "\t" + pc.fracproddict[inst].print());
            }
        }

        Dictionary<string, Dictionary<string, int>> specialinputdict = new Dictionary<string, Dictionary<string, int>>();
        private void read_special_input(string inst)
        {
            openFileDialog1.InitialDirectory = Form1.folder;
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.Title = "Select special input file for " + inst;
            Console.WriteLine("opendialog1.Show:");
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                memo("Reading special " + openFileDialog1.FileName);
                int nchange = 0;
                using (StreamReader sr = new StreamReader(openFileDialog1.FileName))
                {
                    string hline = sr.ReadLine();
                    string[] hwords = hline.Split('\t');
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        string[] w = line.Split('\t');
                        var pc = Form1.findprogram(w[0]);
                        if (pc != null && w.Length > 1)
                        {
                            for (int i=1;i<w.Length;i++)
                            {
                                if (!String.IsNullOrEmpty(w[i]))
                                {
                                    int n = util.tryconvert(w[i]);
                                    if (n >= 0)
                                    {
                                        string sem = hwords[i].Replace("T1 ", "");
                                        var bc = pc.getbatch(sem);
                                        if (bc != null)
                                        {
                                            bc.actualsemstud[1] = n;
                                        }
                                        if (!specialinputdict.ContainsKey(pc.name))
                                        {
                                            specialinputdict.Add(pc.name, new Dictionary<string, int>());
                                        }
                                        if (!specialinputdict[pc.name].ContainsKey(hwords[i]))
                                        {
                                            specialinputdict[pc.name].Add(hwords[i], n);
                                        }
                                        else
                                        {
                                            specialinputdict[pc.name][hwords[i]] = n;
                                        }
                                        nchange++;

                                        //else
                                        //{
                                        //    double?[] actualstud = new double?[programbatchclass.maxsem];
                                        //    for (int j = 0; j < programbatchclass.maxsem; j++)
                                        //    {
                                        //        actualstud[j] = null;
                                        //    }
                                        //    actualstud[1] = n;
                                        //    programbatchclass b = new programbatchclass(actualstud, pc.id, sem);
                                        //    pc.batchlist.Add(b);
                                        //}
                                    }
                                }
                            }
                        }

                    }
                }
                memo("Changed values: " + nchange);
            }
        }

        private void Excelbutton_Click(object sender, EventArgs e)
        {
            prognos = false;
            lastsemwithdata = TBlastsem.Text;
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();

            string folder = util.timestampfolder(@"C:\Temp\Excel planning sheets per institution");
            //string folder = util.timestampfolder(Form1.folder + @"\Excel planning sheets per institution");
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            int startyear = 2023;
            int endyear = 2030;
            int sy = util.tryconvert(TBstartyear.Text);
            if (sy > 0)
                startyear = sy;
            int ey = util.tryconvert(TBendyear.Text);
            if (ey > 0)
                endyear = ey;
            string startsem = "VT" + startyear % 100;
            string endsem = "HT" + endyear % 100;

            Dictionary<string, string> fninst = new Dictionary<string, string>();
            Dictionary<string, Excel.Workbook> xldict = new Dictionary<string, Excel.Workbook>();
            Dictionary<string, Dictionary<string, Excel.Worksheet>> sheetdictdict = new Dictionary<string, Dictionary<string, Excel.Worksheet>>();

            List<string> instlist = Form1.institutiondict.Keys.ToList();
            instlist.Add(Form1.hda);

            foreach (string inst in instlist)
            {
                fninst.Add(inst, util.unusedfn(folder + "HST-planering " + Form1.instshortdict[inst] +" "+util.yymmdd()+" "+DateTime.Now.ToString("hh.mm")+" .xlsx"));
                Excel.Workbook xl = xlApp.Workbooks.Add();
                xldict.Add(inst, xl);
                sheetdictdict.Add(inst, new Dictionary<string, Excel.Worksheet>());
            }



            int ncat = 0;
            int maxcount = 333333;

            List<string> sheetnames = new List<string>();

            foreach (string inst in fninst.Keys)
            {
                rrow.Clear();
                batsemref.Clear();
                progkurssheetnames.Clear();
                memo(inst);
                //if (inst != "Institutionen för information och teknik")
                //    continue;
                //if (!CBinst[inst].Checked)
                //    continue;
                if (!LBinst.SelectedItems.Contains(inst))
                    continue;

                if (CBspecialinput.Checked)
                {
                    read_special_input(inst);
                }

                List<programclass> qprog;
                IEnumerable<programclass> qpr;
                if (RB_homeinst.Checked)
                    qpr = from c in Form1.origprogramdict
                             where c.Value.utype != "Kurspaket"
                             where c.Value.homeinst == inst select c.Value;
                else
                    qpr = from c in Form1.origprogramdict
                             where c.Value.utype != "Kurspaket"
                             where c.Value.fracproddict.ContainsKey(inst) || c.Value.homeinst == inst
                             select c.Value;
                if (CBsortprog.Checked)
                {
                    qprog = qpr.OrderBy(c => (c.homeinst!=inst)).ThenBy(c=>!c.hasbeginners(startsem,endsem)).ThenBy(c=>c.name).ToList();
                }
                else
                {
                    qprog = qpr.ToList();
                }

                var qpaket = (from c in Form1.origprogramdict
                         where c.Value.utype == "Kurspaket"
                         where c.Value.homeinst == inst || inst == Form1.hda
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
                    case "Institutionen för hälsa och välfärd":
                        var qhv6 = (from c in qprog where c.semesters == 6 select c).ToList();
                        if (qhv6.Count > 0)
                        {
                            programclass p6 = programclass.clone(qhv6);
                            p6.name = "Nytt program 180 hp";
                            qprog.Add(p6);
                        }
                        var qhvm = (from c in qprog where c.name.StartsWith("Spec") select c).ToList();
                        if (qhvm.Count > 0)
                        {
                            programclass pm = programclass.clone(qhvm);
                            pm.name = "Nytt spec-ssk-program";
                            qprog.Add(pm);
                        }
                        break;
                        
                    case "Institutionen för språk, litteratur och lärande":
                        var qs6 = (from c in qprog where c.semesters == 6 select c).ToList();
                        if (qs6.Count > 0)
                        {
                            programclass p6 = programclass.clone(qs6);
                            p6.name = "Nytt program 180 hp";
                            qprog.Add(p6);
                        }
                        var qsm = (from c in qprog where c.semesters == 4 where c.is_advanced() select c).ToList();
                        if (qsm.Count > 0)
                        {
                            programclass pm = programclass.clone(qsm);
                            pm.name = "Nytt masterprogram 120 hp";
                            qprog.Add(pm);
                        }
                        break;
                    case "Institutionen för lärarutbildning":
                        var ql8 = (from c in qprog where c.semesters == 8 select c).ToList();
                        if (ql8.Count > 0)
                        {
                            programclass p6 = programclass.clone(ql8);
                            p6.name = "Nytt program 240 hp";
                            qprog.Add(p6);
                        }
                        var qlm = (from c in qprog where c.semesters == 4 where c.is_advanced() select c).ToList();
                        if (qlm.Count > 0)
                        {
                            programclass pm = programclass.clone(qlm);
                            pm.name = "Nytt masterprogram 120 hp";
                            qprog.Add(pm);
                        }
                        break;
                    case "Institutionen för kultur och samhälle":
                        var qks6 = (from c in qprog where c.semesters == 6 select c).ToList();
                        //if (qks6.Count > 0)
                        //{
                        //    programclass p6 = programclass.clone(qks6);
                        //    p6.name = "Filmproduktion 180 hp";
                        //    qprog.Add(p6);
                        //}
                        //if (qks6.Count > 0)
                        //{
                        //    programclass p6 = programclass.clone(qks6);
                        //    p6.name = "Musik- och ljudproduktion 180 hp";
                        //    qprog.Add(p6);
                        //}
                        //if (qks6.Count > 0)
                        //{
                        //    programclass p6 = programclass.clone(qks6);
                        //    p6.name = "Medieproduktion för TV och sociala medier 180 hp";
                        //    qprog.Add(p6);
                        //}
                        //if (qks6.Count > 0)
                        //{
                        //    programclass p6 = programclass.clone(qks6);
                        //    p6.name = "Manus och producentskap 180 hp";
                        //    qprog.Add(p6);
                        //}
                        if (qks6.Count > 0)
                        {
                            programclass p6 = programclass.clone(qks6);
                            p6.name = "Nytt program D 180 hp";
                            qprog.Add(p6);
                        }
                        var qks4 = (from c in qprog where c.semesters == 4 where !c.is_advanced() select c).ToList();
                        if (qks4.Count > 0)
                        {
                            programclass p4 = programclass.clone(qks4);
                            p4.name = "Nytt program 120 hp";
                            qprog.Add(p4);
                        }
                        var qksm = (from c in qprog where c.semesters == 4 where c.is_advanced() select c).ToList();
                        if (qksm.Count > 0)
                        {
                            programclass pm = programclass.clone(qksm);
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
                           where c.Value.homeinst == inst || inst == Form1.hda
                           select c.Value).OrderBy(c=>c.subjectcode).ThenBy(c=>c.name).ToList();
                memo("# courses = " + qcourse.Count());
                int newcourses = 20;
                for (int i = 0; i < newcourses; i++)
                {
                    programclass nc = new programclass(newcoursename + (i + 1) + " (fyll i hp och FK-stud)");
                    qcourse.Add(nc);
                }
                memo("# courses (with new) = " + qcourse.Count());



                Excel.Worksheet mainsheet = xldict[inst].Sheets.Add();
                mainsheet.Name = mainsheetname;
                memo(mainsheet.Name);
                sheetdictdict[inst].Add(mainsheet.Name, mainsheet);

                if (!rrow.ContainsKey(mainsheet.Name))
                    rrow.Add(mainsheet.Name, ProgramNames(mainsheet, qprog, progoffset));

                Excel.Worksheet coursesheet = xldict[inst].Sheets.Add();
                coursesheet.Name = coursesheetname;
                memo(coursesheet.Name);
                sheetdictdict[inst].Add(coursesheet.Name, coursesheet);

                Excel.Worksheet retsheet = xldict[inst].Sheets.Add();
                retsheet.Name = retsheetname;
                memo(retsheet.Name);
                sheetdictdict[inst].Add(retsheet.Name, retsheet);
                RetentionSheet(retsheet, qprog, allmaxsem);

                Excel.Worksheet detailsheet = xldict[inst].Sheets.Add();
                detailsheet.Name = detailsheetname;
                memo(detailsheet.Name);
                sheetdictdict[inst].Add(detailsheet.Name, detailsheet);

                Dictionary<int, Excel.Worksheet> progkurssheets = new Dictionary<int, Excel.Worksheet>();
                for (int year = startyear; year <= endyear; year++)
                {
                    progkurssheetnames.Add(year, progkurssheetname + year);
                    Excel.Worksheet progkurssheet = xldict[inst].Sheets.Add();
                    progkurssheet.Name = progkurssheetnames[year];
                    memo(progkurssheet.Name);
                    sheetdictdict[inst].Add(progkurssheet.Name, progkurssheet);
                    progkurssheets.Add(year, progkurssheet);
                }



                fill_coursehd(coursesheet, qprog, qpaket, qcourse, startsem, endsem);

                fill_planhd(mainsheet,detailsheet,qprog,startsem, endsem);

                Excel.Worksheet batsheet = xldict[inst].Sheets.Add();
                batsheet.Name = batsheetname;

                memo(batsheet.Name);
                sheetdictdict[inst].Add(batsheet.Name, batsheet);
                BatchSheet(batsheet, qprog, startsem, endsem);

                for (int year = startyear; year <= endyear; year++)
                {
                    memo(progkurssheets[year].Name);
                    ProgkursSheet(progkurssheets[year], qprog, qcourse, year, inst);
                }

                memo(mainsheet.Name);
                PlanSheet(mainsheet, qprog, startsem, endsem,inst);

                memo(detailsheet.Name);
                DetailSheet(detailsheet, qprog, startsem, endsem,inst);

                Excel.Worksheet paketsheet = xldict[inst].Sheets.Add();
                paketsheet.Name = paketsheetname;
                memo(paketsheet.Name);
                PaketSheet(paketsheet, qpaket, startsem, endsem, inst);

                memo(coursesheet.Name);
                CourseSheet(coursesheet, qprog, qpaket, qcourse, startsem, endsem);

                memo(mainsheet.Name);
                PlanSheet_FKrows(mainsheet, qprog, startsem, endsem, inst);

                if (CBtriangel.Checked)
                {
                    Excel.Worksheet triangelsheet = xldict[inst].Sheets.Add();
                    triangelsheet.Name = triangelsheetname;
                    memo(triangelsheet.Name);
                    TriangelSheet(triangelsheet, startsem, endsem);
                }

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
                //break;
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

        private void ExcelForm_Load(object sender, EventArgs e)
        {
        }



        private void convertonefile(string fn, Excel.Application xlApp, Excel.XlFileFormat format)
        {
            Excel.Workbook xl = xlApp.Workbooks.Open(fn);
            string fntext = fn.Replace(".xlsx", ".txt");
            memo("Saving to " + fntext);
            xl.SaveAs(fntext, format);
            xl.Close();
            Marshal.ReleaseComObject(xl);
        }

        private void convertfolderbutton_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();

            //openFileDialog1.InitialDirectory = docfolder;
            //openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            folderBrowserDialog1.Description = "Select folder with files to convert";

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string folder = folderBrowserDialog1.SelectedPath;
                if (!Directory.Exists(folder))
                {
                    memo(folder + " not found");
                    return;
                }
                if (Directory.GetFiles(folder).Count() == 0)
                {
                    memo("No files in " + folder);
                    return;
                }
                memo("Reading files from " + folder);

                foreach (string fn in Directory.GetFiles(folder))
                {
                    if (!fn.Contains(".xlsx"))
                        continue;
                    convertonefile(fn, xlApp, Excel.XlFileFormat.xlUnicodeText);
                }

            }
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

        private void coursecheckbutton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            folderBrowserDialog1.Description = "Select folder with course files to check against";

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string folder = folderBrowserDialog1.SelectedPath;
                if (!Directory.Exists(folder))
                {
                    memo(folder + " not found");
                    return;
                }
                if (Directory.GetFiles(folder).Count() == 0)
                {
                    memo("No files in " + folder);
                    return;
                }
                memo("Reading files from " + folder);

                foreach (string fn in Directory.GetFiles(folder))
                {
                    if (!fn.Contains(".txt"))
                        continue;
                    using (StreamReader sr = new StreamReader(fn))
                    {
                        string hline = sr.ReadLine();
                        if (hline.StartsWith("Skapade kurstillfällen"))
                            hline = sr.ReadLine();
                        string[] hwords = hline.Split('\t');
                        int hname = -1;
                        int hcode = -1;
                        for (int i = 0; i < hwords.Length; i++)
                        {
                            if (hwords[i] == "Benämning")
                                hname = i;
                            if (hwords[i] == "FRISTÅENDE KURSER")
                                hname = i;
                            if (hwords[i] == "Programkurser")
                                hname = i;
                            if (hwords[i] == "Kod")
                                hcode = i;
                        }
                        if (hname < 0 || hcode < 0)
                        {
                            memo(hline+"\t"+fn);
                            continue;
                        }
                        int hneeded = Math.Max(hname, hcode);
                        while (!sr.EndOfStream)
                        {
                            string line = sr.ReadLine();
                            string[] words = line.Split('\t');
                            if (words.Length - 1 < hneeded)
                                continue;
                            string code = words[hcode];
                            string name = words[hname].Trim('"');
                            if (String.IsNullOrEmpty(code))
                                continue;
                            if (String.IsNullOrEmpty(name))
                                continue;
                            programclass fk = Form1.findcourse(code, Form1.fkdict, Form1.fkcodedict);
                            if (fk == null)
                                fk = Form1.findcourse(name, Form1.fkdict, Form1.fkcodedict);
                            if (fk == null && code.StartsWith("K22"))
                                fk = Form1.findcourse(name, Form1.paketdict, Form1.fkcodedict);
                            if (fk == null)
                            {
                                memo("Not found\t" + code + "\t" + name);
                            }
                        }
                    }

                }

            }

        }

        private void read_planeringstal()
        {
            string fn = Form1.folder + @"\planeringstal"+DateTime.Now.Year+".txt";
            if (!File.Exists(fn))
            {
                openFileDialog1.Title = "Välj fil med planeringstal";
                //Console.WriteLine("opendialog1.Show:");
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fn = openFileDialog1.FileName;
                }
                else
                    return;
            }

            using (StreamReader sr = new StreamReader(fn))
            { 
                string header = sr.ReadLine();
                string[] hwords = header.Split('\t');
                sr.ReadLine();

                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    string[] words = line.Split('\t');
                    string name = words[0];

                    programclass pc = Form1.findprogram(name);

                    if (pc == null)
                    {
                        memo(name + " not found");
                        continue;
                    }

                    for (int i = 1; i < hwords.Length; i++)
                    {
                        if (i < words.Length)
                        {
                            int n = util.tryconvert(words[i]);
                            programbatchclass bc = pc.getbatch(hwords[i]);
                            if (bc != null)
                                bc.budget_T1 = n;
                        }
                    }
                }
            }
        }

        private void print_fracproddict(List<programclass> q)
        {
            string s = "";
            foreach (string inst in Form1.instshortdict.Keys)
                s += "\t" + Form1.instshortdict[inst];
            memo(s);
            memo("hstpeng");
            foreach (programclass pc in q)
            {
                string s2 = pc.name;
                foreach (string inst in Form1.instshortdict.Keys)
                {
                    if (pc.fracproddict.ContainsKey(inst))
                        s2 += "\t" + pc.fracproddict[inst].hstpeng;
                    else
                        s2 += "\t";
                }
                memo(s2);
            }
            memo(s);
            memo("frachst");
            foreach (programclass pc in q)
            {
                string s2 = pc.name;
                foreach (string inst in Form1.instshortdict.Keys)
                {
                    if (pc.fracproddict.ContainsKey(inst))
                        s2 += "\t" + pc.fracproddict[inst].frachst;
                    else
                        s2 += "\t";
                }
                memo(s2);
            }
            memo(s);
            memo("frachstmoney");
            foreach (programclass pc in q)
            {
                string s2 = pc.name;
                foreach (string inst in Form1.instshortdict.Keys)
                {
                    if (pc.fracproddict.ContainsKey(inst))
                        s2 += "\t" + pc.fracproddict[inst].frachstmoney;
                    else
                        s2 += "\t";
                }
                memo(s2);
            }
            memo(s);
            memo("hprpeng");
            foreach (programclass pc in q)
            {
                string s2 = pc.name;
                foreach (string inst in Form1.instshortdict.Keys)
                {
                    if (pc.fracproddict.ContainsKey(inst))
                        s2 += "\t" + pc.fracproddict[inst].hprpeng;
                    else
                        s2 += "\t";
                }
                memo(s2);
            }
            memo(s);
            memo("frachpr");
            foreach (programclass pc in q)
            {
                string s2 = pc.name;
                foreach (string inst in Form1.instshortdict.Keys)
                {
                    if (pc.fracproddict.ContainsKey(inst))
                        s2 += "\t" + pc.fracproddict[inst].frachpr;
                    else
                        s2 += "\t";
                }
                memo(s2);
            }
        }

        private void Prognosbutton_Click(object sender, EventArgs e)
        {
            prognos = true;

            lastsemwithdata = TBlastsem.Text;
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();

            read_planeringstal();

            string folder = util.timestampfolder(@"C:\Temp\Excel forecast sheets");
            //string folder = util.timestampfolder(Form1.folder + @"\Excel planning sheets per institution");
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);


            //Dictionary<string, string> fninst = new Dictionary<string, string>();
            //Dictionary<string, Excel.Workbook> xldict = new Dictionary<string, Excel.Workbook>();
            //Dictionary<string, Dictionary<string, Excel.Worksheet>> sheetdictdict = new Dictionary<string, Dictionary<string, Excel.Worksheet>>();
            Dictionary<string, Excel.Worksheet> sheetdict = new Dictionary<string, Excel.Worksheet>();

            //foreach (string inst in Form1.institutiondict.Keys)
            //{
            //    fninst.Add(inst, util.unusedfn(folder + "HST-planering " + Form1.instshortdict[inst] + " " + util.yymmdd() + ".xlsx"));
            //    Excel.Workbook xl = xlApp.Workbooks.Add();
            //    xldict.Add(inst, xl);
            //    sheetdictdict.Add(inst, new Dictionary<string, Excel.Worksheet>());
            //}

            string fn = util.unusedfn(folder + "HST-prognos HDa " + util.yymmdd() + " " + DateTime.Now.ToString("hh.mm") + " .xlsx");
            Excel.Workbook xl = xlApp.Workbooks.Add();


            int ncat = 0;
            int maxcount = 333333;

            List<string> sheetnames = new List<string>();


            List<programclass> qprog = (from c in Form1.origprogramdict
                            where c.Value.utype != "Kurspaket"
                            where c.Value.homeinst != Form1.utaninst
                            select c.Value).ToList();

            //print_fracproddict(qprog);
            //return;

            var qpaket = (from c in Form1.origprogramdict
                            where c.Value.utype == "Kurspaket"
                          where c.Value.homeinst != Form1.utaninst
                          select c.Value).ToList();

            int nprog = qprog.Count;
            int allmaxsem = (from c in qprog select c.semesters).Max();

            memo("nprog = " + nprog);

            List<programclass> qcourse;
            qcourse = (from c in Form1.fkdict
                        where c.Value.activecourse
                        select c.Value).OrderBy(c => c.subjectcode).ThenBy(c => c.name).ToList();
            memo("# courses = " + qcourse.Count());
            int newcourses = 20;
            for (int i = 0; i < newcourses; i++)
            {
                programclass nc = new programclass(newcoursename + (i + 1) + " (fyll i hp och FK-stud)");
                qcourse.Add(nc);
            }
            memo("# courses (with new) = " + qcourse.Count());


            string startsem = "VT23";
            string endsem = "HT24";

            string instinst = Form1.hda;

            Excel.Worksheet mainsheet = xl.Sheets.Add();
            mainsheet.Name = mainsheetname;
            memo(mainsheet.Name);
            sheetdict.Add(mainsheet.Name, mainsheet);

            if (!rrow.ContainsKey(mainsheet.Name))
                rrow.Add(mainsheet.Name, ProgramNames(mainsheet, qprog, progoffset));

            Excel.Worksheet paketsheet = xl.Sheets.Add();
            paketsheet.Name = paketsheetname;
            if (!rrow.ContainsKey(paketsheet.Name))
                rrow.Add(paketsheet.Name, ProgramNames(paketsheet, qpaket, progoffset));

            Excel.Worksheet coursesheet = xl.Sheets.Add();
            coursesheet.Name = coursesheetname;
            if (!rrow.ContainsKey(coursesheet.Name))
                rrow.Add(coursesheet.Name, ProgramNames(coursesheet, qcourse, progoffset));

            Excel.Worksheet retsheet = xl.Sheets.Add();
            retsheet.Name = retsheetname;
            memo(retsheet.Name);
            sheetdict.Add(retsheet.Name, retsheet);
            RetentionSheet(retsheet, qprog, allmaxsem);

            Excel.Worksheet retcoursesheet = xl.Sheets.Add();
            retcoursesheet.Name = retcoursesheetname;
            memo(retcoursesheet.Name);
            sheetdict.Add(retcoursesheet.Name, retcoursesheet);
            RetentionSheet(retcoursesheet, qcourse, allmaxsem);

            Excel.Worksheet retpaketsheet = xl.Sheets.Add();
            retpaketsheet.Name = retpaketsheetname;
            memo(retpaketsheet.Name);
            sheetdict.Add(retpaketsheet.Name, retpaketsheet);
            RetentionSheet(retpaketsheet, qpaket, allmaxsem);

            Excel.Worksheet detailsheet = xl.Sheets.Add();
            detailsheet.Name = detailsheetname;
            memo(detailsheet.Name);
            sheetdict.Add(detailsheet.Name, detailsheet);
            if (!rrow.ContainsKey(detailsheet.Name))
                rrow.Add(detailsheet.Name, ProgramNames(detailsheet, qprog, progoffset));

            fill_planhd_prognos(mainsheet, detailsheet, qprog, startsem, endsem);

            Excel.Worksheet batsheet = xl.Sheets.Add();
            batsheet.Name = batsheetname;

            memo(batsheet.Name);
            sheetdict.Add(batsheet.Name, batsheet);
            BatchSheet(batsheet, qprog, startsem, endsem);


            memo(mainsheet.Name);
            PlanSheetPrognos(mainsheet, qprog, startsem, endsem, instinst);

            memo(detailsheet.Name);
            DetailSheet(detailsheet, qprog, startsem, endsem, instinst);

            memo(paketsheet.Name);
            PaketSheetPrognos(paketsheet, qpaket, startsem, endsem, instinst);

            memo(coursesheet.Name);
            CourseSheetPrognos(coursesheet, qprog, qpaket, qcourse, startsem, endsem);

            memo(mainsheet.Name);
            PlanSheet_FKrows(mainsheet, qprog, startsem, endsem, instinst);

            foreach (string inst in Form1.institutiondict.Keys)
            {
                Excel.Worksheet instsheet = xl.Sheets.Add();
                instsheet.Name = Form1.instshortdict[inst]+" "+mainsheetname;
                memo(instsheet.Name);
                sheetdict.Add(instsheet.Name, instsheet);
                var qpinst = (from c in Form1.origprogramdict
                         where c.Value.utype != "Kurspaket"
                         where c.Value.fracproddict.ContainsKey(inst)
                         select c.Value).ToList();
                SheetWithHeader(instsheet, qprog.Count + progoffset, planhd);
                PlanSheetPrognos(instsheet, qpinst, startsem, endsem, inst);
                PlanSheet_FKrows(instsheet, qpinst, startsem, endsem, inst);
            }

            Excel.Worksheet sumsheet = xl.Sheets.Add();
            sumsheet.Name = sumsheetname;
            memo(sumsheet.Name);
            sheetdict.Add(sumsheet.Name, sumsheet);

            SumSheet(sumsheet, qprog, startsem, endsem);

            //mainsheet.Select();

            memo("Saving to " + fn);
            xl.SaveAs(fn);

            foreach (string sc in sheetdict.Keys)
            {
                Marshal.ReleaseComObject(sheetdict[sc]);
            }
            xl.Close();
            Marshal.ReleaseComObject(xl);

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

        public string clean_utbnamn(string un)
        {
            int maxlen = 40;
            if (un.Length <= maxlen)
                return un.Trim();
            return un.Substring(0, maxlen).Trim();
        }

        List<CSNpostclass> CSNlist = new List<CSNpostclass>();
        private void omstallningbutton_Click(object sender, EventArgs e)
        {
            Excel.Application oXL;
            //Excel.Workbook oWB;

            oXL = new Excel.Application();

            string username = Environment.UserName;
            memo(username);
            string fnapp = @"C:\Users\"+username+@"\OneDrive - Högskolan Dalarna\Dokument\Närvärld\UKÄ - Omställningsstudiestöd, antal sökande, sökt lärosäte och utbildning 2024-10-08.xlsx";
            string fnacc = @"C:\Users\"+username+@"\OneDrive - Högskolan Dalarna\Dokument\Närvärld\UKÄ - Omställningsstudiestöd, antal beviljade personer med skola och utbildningar 2024-10-08.xlsx";

            oXL.Workbooks.Open(fnapp);
            oXL.Workbooks.Open(fnacc);

            foreach (var wb in oXL.Workbooks)
            {
                memo(((Excel.Workbook)wb).Title);
            }
            foreach (Excel.Worksheet ws in oXL.Workbooks[1].Sheets)
            {
                string uni = ws.Name.Trim('_');
                memo(uni);
                int n = 0;
                for (int i = 5; i < 400; i++)
                {
                    string s = ws.Cells[i, 1].Value.ToString();
                    if (s.StartsWith("Total"))
                        break;
                    if (!s.StartsWith("Ansök"))
                        continue;
                    n++;
                    string sok = s.Replace("Ansökanstillfälle ", "").Replace("Ansökningstillfälle ", "");
                    string kod = ws.Cells[i, 3].Value.ToString();
                    string utbnamn = clean_utbnamn(ws.Cells[i, 4].Value.ToString());
                    int nstud = Convert.ToInt32(ws.Cells[i, 5].Value);

                    CSNpostclass pc = new CSNpostclass();
                    pc.uni = uni;
                    pc.appldate = sok;
                    pc.name = utbnamn;
                    pc.applicants = nstud;
                    pc.code = kod;

                    CSNlist.Add(pc);
                }
                memo(n.ToString());
            }

            foreach (Excel.Worksheet ws in oXL.Workbooks[2].Sheets)
            {
                string uni = ws.Name.Trim('_');
                memo(uni);
                int n = 0;
                int nnf = 0;
                for (int i = 5; i < 400; i++)
                {
                    string s = ws.Cells[i, 1].Value.ToString();
                    if (s.StartsWith("Total"))
                        break;
                    if (!s.StartsWith("Ansök"))
                        continue;
                    n++;
                    string sok = s.Replace("Ansökanstillfälle ", "").Replace("Ansökningstillfälle ", "");
                    //if (sok.EndsWith("2022"))
                    //    continue;
                    string kod = ws.Cells[i, 3].Value.ToString();
                    string utbnamn = clean_utbnamn(ws.Cells[i, 4].Value.ToString());
                    int nstud = Convert.ToInt32(ws.Cells[i, 5].Value);

                    var qc = from c in CSNlist
                             where c.uni == uni
                             where c.appldate == sok
                             where c.name == utbnamn
                             where c.code == kod
                             select c;
                    if (qc.Count() == 1)
                    {
                        CSNpostclass pc = qc.First();
                        pc.accepted = nstud;
                    }
                    else
                    {
                        CSNpostclass pc = new CSNpostclass();
                        pc.uni = uni;
                        pc.appldate = sok;
                        pc.name = utbnamn;
                        pc.accepted = nstud;
                        pc.code = kod;

                        CSNlist.Add(pc);

                        nnf++;
                    }


                }
                memo(n.ToString());
                memo("not found " + nnf);
            }

            foreach (Excel.Workbook wb in oXL.Workbooks)
            {
                memo(((Excel.Workbook)wb).Title);
                wb.Close();
                Marshal.ReleaseComObject(wb);
            }
            //quit and release
            oXL.Quit();
            Marshal.ReleaseComObject(oXL);

            memo("appldate");
            foreach (string s in (from c in CSNlist select c.appldate).Distinct())
                memo(s);

            memo("code");
            foreach (string s in (from c in CSNlist select c.code).Distinct())
                memo(s);

            memo("name");
            foreach (string s in (from c in CSNlist select c.name).Distinct())
                memo(s);

        }
    }
}
