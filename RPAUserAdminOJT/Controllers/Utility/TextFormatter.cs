using System;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace RPAUserAdminOJT.Controllers.Utility
{
    public class TextFormatter
    {
        public static string Clean(string text)
        {
            if (!String.IsNullOrEmpty(text))
            {
                string result = string.Concat(Regex.Replace(text, @"(?i)[\p{L}-[ña-z]]+", m => m.Value.Normalize(NormalizationForm.FormD))
                .Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark));
                result = result.Replace("ñ", "n");
                result = result.Replace("Ñ", "n");
                result = result.Replace(" ", "");
                result = Regex.Replace(result, @"[^\w\s!.]", "");
                return result.ToLower();
            }
            return null;
        }

        public static string FirstLetterToUpperCase(string word)
        {
            if (word == null || word.Length <= 0)
                return null;

            if (word.Length > 1)
                return char.ToUpper(word[0]) + word.Substring(1);

            return word.ToUpper();
        }
    }
}
