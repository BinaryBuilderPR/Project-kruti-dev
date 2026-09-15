using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace KrutiDevWordAddIn
{
    /// <summary>
    /// Pure C# Kruti Dev 010 <-> Unicode Devanagari Bidirectional Converter for MS Word VSTO Add-in.
    /// </summary>
    public static class KrutiDevConverter
    {
        private static readonly Dictionary<string, string> MultiMap = new Dictionary<string, string>
        {
            {"vksS", "औ"}, {"vkS", "औ"}, {"vks", "ओ"}, {"vk", "आ"}, {"v", "अ"},
            {"bZ", "ई"}, {"b", "इ"}, {"Å", "ऊ"}, {"m", "उ"}, {"_", "ऋ"}, {",s", "ऐ"}, {",", "ए"},
            {"Dr", "क्त"}, {"iz", "प्र"}, {"nz", "द्र"}, {"Ø", "क्र"}, {"Vª", "ट्र"}, {"Mª", "ड्र"},
            {"Í", "द्द"}, {"Î", "द्ध"}, {"Ï", "द्घ"}, {"Ð", "द्य"}, {"Ñ", "कृ"}, {"Ò", "दृ"},
            {"Ó", "हृ"}, {"®", "रु"}, {"¯", "रू"}, {"}", "द्व"}, {"|", "द्य"},
            {"[kk", "खा"}, {"[k", "ख"}, {"?k", "घ"}, {"Fk", "थ"}, {"/k", "ध"}, {"Hk", "भ"},
            {"'k", "श"}, {"\"k", "ष"}, {"{k", "क्ष"}, {".k", "ण"},
            {"=", "त्र"}, {"K", "ज्ञ"}, {"J", "श्र"}, {"Wa", "ँ"}, {"Wk", "ँ"}, {"W", "ँ"}
        };

        private static readonly Dictionary<char, string> SingleMap = new Dictionary<char, string>
        {
            {'d', "क"}, {'x', "ग"}, {'p', "च"}, {'N', "छ"}, {'t', "ज"}, {'T', "झ"}, {'¥', "ञ"},
            {'V', "ट"}, {'B', "ठ"}, {'M', "ड"}, {'<', "ढ"}, {'r', "त"}, {'n', "द"}, {'u', "न"},
            {'i', "प"}, {'Q', "फ"}, {'c', "ब"}, {'e', "म"}, {';', "य"}, {'j', "र"}, {'y', "ल"},
            {'o', "व"}, {'l', "स"}, {'g', "ह"},
            {'D', "क्"}, {'X', "ग्"}, {'P', "च्"}, {'Y', "ल्"}, {'R', "त्"}, {'F', "थ्"},
            {'/', "ध्"}, {'U', "न्"}, {'I', "प्"}, {'C', "ब्"}, {'H', "भ्"}, {'E', "म्"},
            {'O', "व्"}, {'L', "स्"}, {'\'', "श्"}, {'"', "ष्"}, {'[', "ख्"}, {'?', "घ्"},
            {'{', "क्ष्"}, {'.', "ण्"}, {'~', "्"},
            {'k', "ा"}, {'h', "ी"}, {'q', "ु"}, {'w', "ू"}, {'`', "ृ"}, {'s', "े"}, {'S', "ै"},
            {'a', "ं"}, {'z', "्र"}, {'µ', "ृ"},
            {'A', "।"}, {']', ","}, {'%', ":"}, {'&', "-"}
        };

        public static string KrutiToUnicode(string krutiText)
        {
            if (string.IsNullOrEmpty(krutiText)) return "";

            string res = krutiText;
            foreach (var kvp in MultiMap)
            {
                res = res.Replace(kvp.Key, kvp.Value);
            }

            // Chhoti-i reordering
            res = Regex.Replace(res, @"f((?:[\u0915-\u0939]्|[A-Z\[\?\{'""])*[\u0915-\u0939d-hjl-pqrt-x=KJ]|Dr|iz|nz|Ø)", m =>
            {
                string cluster = m.Groups[1].Value;
                StringBuilder sb = new StringBuilder();
                foreach (char ch in cluster)
                {
                    sb.Append(SingleMap.ContainsKey(ch) ? SingleMap[ch] : ch.ToString());
                }
                return sb.ToString() + "ि";
            });

            // Single mappings
            StringBuilder finalSb = new StringBuilder();
            foreach (char ch in res)
            {
                finalSb.Append(SingleMap.ContainsKey(ch) ? SingleMap[ch] : ch.ToString());
            }
            res = finalSb.ToString();

            // Reph (Z) reordering
            res = Regex.Replace(res, @"((?:[\u0915-\u0939]्)*[\u0915-\u0939][\u093E-\u094C\u0901-\u0903]*)Z", "र्$1");

            return res;
        }
    }
}
