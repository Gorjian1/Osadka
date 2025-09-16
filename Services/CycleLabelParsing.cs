using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Osadka.Services
{
    public static class CycleLabelParsing
    {
        public static string? ExtractDateTail(string? label)
        {
            if (string.IsNullOrWhiteSpace(label)) return null;

            var s = label.TrimEnd();
            int i = s.Length - 1;
            while (i >= 0 && !char.IsWhiteSpace(s[i])) i--;

            if (i < 0) return s;
            return s.Substring(i + 1);
        }
    }
}
