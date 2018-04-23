using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CefSharp.MinimalExample.WinForms
{
    public class ExcelRecord
    {
        public long RowNumber { get; set; }
        public string Refl { get; set; }
        public string Vorsible { get; set; }
        public string Verb { get; set; }
        public string Infinitiv { get; set; }
        public string Partizip { get; set; }
        public string Perfekt { get; set; }
        public string Präsens { get; set; }
        public string Präteritum { get; set; }
        public string PP { get; set; }

        public bool Flag { get; set; } = false;

        public override string ToString()
        {
            return $"{RowNumber} {Flag}: {Refl} {Vorsible} {Verb}, Perfekt = {Perfekt} {Partizip}, Präsens = {Präsens}, Präteritum = {Präteritum}";
        }

        public string FinalTrembalVerb
        {
            get
            {
                return Vorsible + Verb;
            }
        }
        public string FinalVerb
        {
            get
            {
                return (string.IsNullOrEmpty(Refl) ? "" : Refl + " ") + Vorsible + Verb;
            }
        }
    }
}
