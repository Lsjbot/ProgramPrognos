using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProgramPrognos
{
    internal class dictclass
    {
        Dictionary<string, string> dict = new Dictionary<string, string>();

        public dictclass(string[] hwords, string[] words)
        {
            for (int i=0; i < Math.Min(hwords.Length,words.Length); i++)
            {
                dict.Add(hwords[i], words[i]);
            }
        }

        public void Add(string key,string value)
        {
            if (!dict.ContainsKey(key))
                dict.Add(key, value);
        }

        public bool Has(string key)
        {
            return dict.ContainsKey(key) && !String.IsNullOrEmpty(dict[key]);
        }
        
        public string Get(string key)
        {
            if (dict.ContainsKey(key))
                return dict[key];
            else
                return null;
        }

        public int Getint(string key)
        {
            int result;
            if (int.TryParse(Get(key), out result))
                return result;
            else
                return 0;
        }

        public double Getdouble(string key)
        {
            return util.tryconvertdouble(Get(key));
        }

        public string Printline(List<string> keys)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string key in keys)
                sb.Append(Get(key) + "\t");
            return sb.ToString();
        }

        
    }
}
