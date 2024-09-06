using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;

namespace ProgramPrognos
{
    public class util
    {
        public static int tryconvert(string word)
        {
            int i = -1;

            if (word.Length == 0)
                return i;

            try
            {
                i = Convert.ToInt32(word);
            }
            catch (OverflowException)
            {
                Console.WriteLine("i Outside the range of the Int32 type: " + word);
            }
            catch (FormatException)
            {
                //if ( !String.IsNullOrEmpty(word))
                //    Console.WriteLine("i Not in a recognizable format: " + word);
            }

            return i;

        }

        public static int tryconvert0(string word)
        {
            int i = 0;

            if (word.Length == 0)
                return i;

            try
            {
                i = Convert.ToInt32(word);
            }
            catch (OverflowException)
            {
                Console.WriteLine("i Outside the range of the Int32 type: " + word);
            }
            catch (FormatException)
            {
                //if ( !String.IsNullOrEmpty(word))
                //    Console.WriteLine("i Not in a recognizable format: " + word);
            }

            return i;

        }

        public static int? tryconvertnull(string word)
        {
            int? i = null;

            if (word.Length == 0)
                return null;

            try
            {
                i = Convert.ToInt32(word);
            }
            catch (OverflowException)
            {
                Console.WriteLine("i Outside the range of the Int32 type: " + word);
            }
            catch (FormatException)
            {
                //if ( !String.IsNullOrEmpty(word))
                //    Console.WriteLine("i Not in a recognizable format: " + word);
            }

            return i;

        }

        public static double tryconvertdouble(string word)
        {
            double i = -1;

            if (word.Length == 0)
                return i;

            try
            {
                i = Convert.ToDouble(word);
            }
            catch (OverflowException)
            {
                Console.WriteLine("i Outside the range of the Double type: " + word);
            }
            catch (FormatException)
            {
                try
                {
                    i = Convert.ToDouble(word.Replace(".", ","));
                }
                catch (FormatException)
                {
                    //Console.WriteLine("i Not in a recognizable double format: " + word.Replace(".", ","));
                }
                //Console.WriteLine("i Not in a recognizable double format: " + word);
            }

            return i;

        }

        public static float tryconvertfloat(string word)
        {
            float i = -1;

            if (word.Length == 0)
                return i;

            try
            {
                i = (float)Convert.ToDouble(word);
            }
            catch (OverflowException)
            {
                Console.WriteLine("i Outside the range of the Double type: " + word);
            }
            catch (FormatException)
            {
                try
                {
                    i = (float)Convert.ToDouble(word.Replace(".", ","));
                }
                catch (FormatException)
                {
                    //Console.WriteLine("i Not in a recognizable double format: " + word.Replace(".", ","));
                }
                //Console.WriteLine("i Not in a recognizable double format: " + word);
            }

            return i;

        }

        public static int qtoint(string q)
        {
            return tryconvert(q.Replace("Q", ""));
        }

        public static int semtoint(string sem)
        {
            return tryconvert(sem.Substring(sem.Length-2, 2));
        }

        public static double SampleGaussian(Random random, double mean, double stddev)
        {
            //From https://gist.github.com/tansey/1444070
            // The method requires sampling from a uniform random of (0,1]
            // but Random.NextDouble() returns a sample of [0,1).
            double x1 = 1 - random.NextDouble();
            double x2 = 1 - random.NextDouble();

            double y1 = Math.Sqrt(-2.0 * Math.Log(x1)) * Math.Cos(2.0 * Math.PI * x2);
            return y1 * stddev + mean;
        }

        public static bool comparesemesters(string sem1, string sem2) //true if sem2 later than sem1
        {
            if (sem1 == sem2)
                return false;
            if (semtoint(sem1) > semtoint(sem2))
            {
                return false;
            }
            else if (semtoint(sem1) < semtoint(sem2))
            {
                return true;
            }
            else
            {
                return (sem1.ToUpper().StartsWith("V"));
            }
        }

        public static int semestercount(string startsem, string sem) //for a batch starting at startsem, which semester (T1 etc) is sem?
        {
            if (sem == startsem)
                return 1;
            else if (comparesemesters(sem, startsem))
                return -1;
            else
            {
                int n = 1;
                string ss = startsem;
                while (ss != sem)
                {
                    n++;
                    ss = util.incrementsemester(ss);
                }
                return n;
            }

        }

        public static string find_batstart(string currentsem,int isem)
            //find the batch that is in T-isem during currentsem
        {
            if (isem == 1)
                return currentsem;
            else
                return shiftsemester(currentsem, 1 - isem);
        }

        public static string incrementsemester(string sem)
        {
            if (sem.StartsWith("VT"))
                return sem.Replace("VT", "HT");
            else
                return "VT" + (util.semtoint(sem) + 1);
        }

        public static string decrementsemester(string sem)
        {
            if (sem.StartsWith("HT"))
                return sem.Replace("HT", "VT");
            else
                return "HT" + (util.semtoint(sem) - 1);
        }

        public static string shiftsemester(string sem, int nsem)
        {
            if ( nsem < 0) //shift pastwards
            {
                int odd = nsem % 2;
                string newsem = sem.Substring(0, 2) + (util.semtoint(sem) + (nsem / 2));
                if (odd == 0)
                    return newsem;
                else
                    return decrementsemester(newsem);

            }
            else //shift futurewards
            {
                int odd = nsem % 2;
                string newsem = sem.Substring(0, 2) + (util.semtoint(sem) + (nsem / 2));
                if (odd == 0)
                    return newsem;
                else
                    return incrementsemester(newsem);

            }
        }

        public static string semester4to2(string sem) //change from VT2021 to VT21
        {
            string s = sem.Substring(0, 2) + semtoint(sem);
            return s;
        }

        public static string semester3to2(string sem) //change from V21 to VT21
        {
            string s = sem.Substring(0, 1) + "T"+ semtoint(sem);
            return s;
        }

        public static int year2to4(int yr)
        {
            if (yr < 100)
                return yr + 2000;
            else
                return yr;

        }

        public static int yearfromsem(string sem)
        {
            return year2to4(semtoint(sem));
        }

        public static string unusedfn(string fnbase)
        {
            string suffix = ".txt";
            if (!fnbase.Contains(suffix))
            {
                suffix = "." + fnbase.Split('.').Last();
            }
            int i = 1;
            string fn = fnbase.Replace(suffix, i + suffix);
            while (File.Exists(fn))
            {
                i++;
                fn = fnbase.Replace(suffix, i + suffix);
            }
            return fn;
        }

        public static string yymmdd()
        {
            DateTime now = DateTime.Now;
            return now.ToString("yyMMdd");//now.Year + "-" + now.Month + "-" + now.Day;
        }

        public static string timestampfolder(string folder)
        {
            return timestampfolder(folder, "");
        }

        public static string timestampfolder(string folder, string prefix)
        {
            DateTime now = DateTime.Now;
            string separator = @"\";
            if (folder.EndsWith(separator))
                separator = "";
            return folder + separator + prefix + now.Year + "-" + now.Month + "-" + now.Day + " " + now.Hour + "-" + now.Minute + @"\";

        }

        public static string hprex = @" \d+([\.\,]\d)? hp";

        public static Tuple<string,double> extract_hp(string name)
        {
            
            foreach (Match m in Regex.Matches(name, hprex))
            {
                double hp = util.tryconvertdouble(m.Value.Trim().Replace(" hp", ""));
                string newname = name.Replace(m.Value, "").Trim(new char[]{ ' ',','});
                return new Tuple<string, double>(newname, hp);
            }
            return new Tuple<string, double>(name, -1);
        }

    }
}
