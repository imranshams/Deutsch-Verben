using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CefSharp.MinimalExample.WinForms
{
    public static class Extentions
    {
        public static ExcelRecord map(this Microsoft.Office.Interop.Excel.Range range, long index)
        {
            return new ExcelRecord
            {
                RowNumber = index,
                Refl = range.Cells[index, "A"].Text,
                Vorsible = range.Cells[index, "B"].Text,
                Verb = range.Cells[index, "C"].Text,
                Infinitiv = range.Cells[index, "D"].Text,
                Partizip = range.Cells[index, "E"].Text,
                Perfekt = range.Cells[index, "F"].Text,
                PP = range.Cells[index, "G"].Text,
                Präsens = range.Cells[index, "J"].Text,
                Präteritum = range.Cells[index, "K"].Text
            };
        }

        public static Boolean HasClass(this HtmlNode element, String className)
        {
            if (element == null) throw new ArgumentNullException(nameof(element));
            if (string.IsNullOrWhiteSpace(className)) throw new ArgumentNullException(nameof(className));
            if (element.NodeType != HtmlNodeType.Element) return false;

            HtmlAttribute classAttrib = element.Attributes["class"];
            if (classAttrib == null) return false;

            Boolean hasClass = CheapClassListContains(classAttrib.Value, className, StringComparison.Ordinal);
            return hasClass;
        }

        /// <summary>Performs optionally-whitespace-padded string search without new string allocations.</summary>
        /// <remarks>A regex might also work, but constructing a new regex every time this method is called would be expensive.</remarks>
        private static Boolean CheapClassListContains(String haystack, String needle, StringComparison comparison)
        {
            if (String.Equals(haystack, needle, comparison)) return true;
            Int32 idx = 0;
            while (idx + needle.Length <= haystack.Length)
            {
                idx = haystack.IndexOf(needle, idx, comparison);
                if (idx == -1) return false;

                Int32 end = idx + needle.Length;

                // Needle must be enclosed in whitespace or be at the start/end of string
                Boolean validStart = idx == 0 || Char.IsWhiteSpace(haystack[idx - 1]);
                Boolean validEnd = end == haystack.Length || Char.IsWhiteSpace(haystack[end]);
                if (validStart && validEnd) return true;

                idx++;
            }
            return false;
        }
    }
}
