using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProgramPrognos
{
    public class fracprodclass
    {
        public double frachst = 0;
        public double frachpr = 0;
        public double frachstmoney = 0;
        public double frachprmoney = 0;
        public double fracmoney = 0;
        public double hstpeng = 0;
        public double hprpeng = 0;
        
        public void add(fracprodclass f2)
        {
            add(f2.frachst, f2.frachpr, f2.frachstmoney, f2.frachprmoney, f2.fracmoney);
        }

        public void add(double hst,double hpr,double hstkr,double hprkr, double kr)
        {
            if (Double.IsNaN(kr))
                Console.WriteLine("Bad number ");
            frachst += hst;
            frachpr += hpr;
            frachstmoney += hstkr;
            frachprmoney += hprkr;
            fracmoney += kr;
            updatepeng();
        }

        public void updatepeng()
        {
            if (frachst != 0)
                hstpeng = frachstmoney / frachst;
            if (frachpr != 0)
                hprpeng = frachprmoney / frachpr;
        }

        public void sumstudents(double nstud, fracprodclass perstud)
        {
            this.frachst += nstud * perstud.frachst;
            this.frachpr += nstud * perstud.frachpr;
            this.frachstmoney += nstud * perstud.frachstmoney;
            this.frachprmoney += nstud * perstud.frachprmoney;
            this.fracmoney += nstud * perstud.fracmoney;
            updatepeng();
        }
        public void normalize(double hstsum, double hprsum, double hstmoneysum, double hprmoneysum, double moneysum)
        {
            frachst = frachst / hstsum;
            frachpr = frachpr / hprsum;
            frachstmoney = frachstmoney / hstmoneysum;
            frachprmoney = frachprmoney / hprmoneysum;
            fracmoney = fracmoney / moneysum;
        }

        public void normalize(fracprodclass norm)
        {
            normalize(norm.frachst, norm.frachpr, norm.frachstmoney, norm.frachprmoney, norm.fracmoney);
        }

        public void normalize(double snorm)
        {
            if (!(snorm > 0))
                return;
            normalize(snorm, snorm, snorm, snorm, snorm);
        }

        public fracprodclass clone()
        {
            fracprodclass fp = new fracprodclass();
            fp.frachst = this.frachst;
            fp.frachpr = this.frachpr;
            fp.frachstmoney = this.frachstmoney;
            fp.frachprmoney = this.frachprmoney;
            fp.fracmoney = this.fracmoney;
            fp.hstpeng = this.hstpeng;
            fp.hprpeng = this.hprpeng;
            return fp;
        }

        public static fracprodclass average(List<fracprodclass> fc)
        {
            fracprodclass fp = new fracprodclass();
            fp.frachst = (from c in fc select c.frachst).Average();
            fp.frachpr = (from c in fc select c.frachpr).Average();
            fp.frachstmoney = (from c in fc select c.frachstmoney).Average();
            fp.frachprmoney = (from c in fc select c.frachprmoney).Average();
            fp.fracmoney = (from c in fc select c.fracmoney).Average();
            fp.hstpeng = (from c in fc select c.hstpeng).Average();
            fp.hprpeng = (from c in fc select c.hprpeng).Average();
            return fp;
        }

        public double prestationsgrad()
        {
            if (this.frachst > 0)
                return this.frachpr / this.frachst;
            else
                return 0.8;
        }

        public static string printheader()
        {
            return "frachst\tfrachpr\tfrachstmoney\tfrachprmoney\tfracmoney\thstpeng\thprpeng";
        }
        public string print()
        {
            StringBuilder sb = new StringBuilder(frachst.ToString("N2"));
            sb.Append("\t" + frachpr.ToString("N2"));
            sb.Append("\t" + frachstmoney.ToString("N2"));
            sb.Append("\t" + frachprmoney.ToString("N2"));
            sb.Append("\t" + fracmoney.ToString("N2"));
            sb.Append("\t" + hstpeng.ToString("N2"));
            sb.Append("\t" + hprpeng.ToString("N2"));

            return sb.ToString();

        }
    }

}
