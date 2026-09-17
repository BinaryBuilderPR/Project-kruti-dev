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
            {"=", "त्र"}, {"K", "ज्ञ"}, {"J", "श्र"}, {"Wa", "ँ"}, {"Wk", "ँ"}, {"W", "ँ"},
            {"&", "-"}
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

        private static readonly Dictionary<string, string> DirectUnicodeToKruti = new Dictionary<string, string>
        {
            {"कार्यालय", "dk;kZy;"}, {"शिक्षा", "f'k{kk"}, {"विभाग", "foHkkx"}, {"झारखण्ड", ">kj[k.M"},
            {"राँची", "jkWaph"}, {"सरायकेला", "ljk;dsyk"}, {"खरसावाँ", "[kjlkokWa"}, {"सरायकेला-खरसावाँ", "ljk;dsyk&[kjlkokWa"},
            {"प्राथमिक", "izkFkfed"}, {"माध्यमिक", "ek/;fed"}, {"उच्च", "mPp"}, {"न्यायालय", "U;k;ky;"},
            {"पदाधिकारी", "inkf/kdkjh"}, {"अधीक्षक", "v/kh{kd"}, {"निदेशक", "funs'kd"}, {"आदेश", "vkns'k"},
            {"ज्ञापांक", "Kkikad"}, {"पत्रांक", "i=kad"}, {"दिनांक", "fnukad"}, {"प्रतिवेदन", "izfrosnu"},
            {"अनुपालन", "vuqikyu"}, {"संचिका", "lafpdk"}, {"भवदीय", "Hkonh;"}, {"हस्ताक्षर", "gLrk{kj"},
            {"विषय", "fo\"k;"}, {"प्रसंग", "izlax"}, {"महोदय", "egksn;"}, {"महाशय", "egk'k;"},
            {"अनुरोध", "vuqjks/k"}, {"आवश्यक", "vko';d"}, {"कार्रवाई", "dk;Zokgh"}, {"कार्यवाही", "dk;Zokgh"},
            {"उपरोक्त", "mijksDr"}, {"संलग्न", "layXu"}, {"वेतन", "osru"}, {"शिक्षक", "f'k{kd"},
            {"शिक्षिका", "f'kf{kdk"}, {"विद्यालय", "fon~;ky;"}, {"प्रधानाध्यापक", "iz/kkuk/;kid"},
            {"प्रभारी", "izHkkjh"}, {"छात्र", "Nk="}, {"छात्रा", "Nk=k"}, {"उपस्थिति", "mifLFkfr"},
            {"अवकाश", "vodk'k"}, {"स्वीकृत", "Lohd`r"}, {"अस्वीकृत", "vLohd`r"}, {"स्पष्टीकरण", "Li\"Vhdj.k"},
            {"शोकॉज", "'kksdkWt"}, {"जाँच", "tkWap"}, {"निलंबन", "fuyacu"}, {"वेतनमान", "osrueku"},
            {"में", "esa"}, {"मैं", "eSa"}, {"की", "dh"}, {"कि", "fd"}, {"के", "ds"}, {"को", "dks"},
            {"से", "ls"}, {"पर", "ij"}, {"का", "dk"}, {"है", "gS"}, {"हैं", "gSa"}, {"था", "Fkk"},
            {"थी", "Fkh"}, {"थे", "Fks"}, {"होगा", "gksxk"}, {"किया", "fd;k"}, {"गया", "x;k"},
            {"गए", "x,"}, {"गई", "xbZ"}, {"जाता", "tkrk"}, {"द्वारा", "}kjk"}, {"तथा", "rFkk"},
            {"एवं", ",oa"}, {"अथवा", "vFkok"}, {"यदि", ";fn"}, {"तो", "rks"}, {"इस", "bl"},
            {"उस", "ml"}, {"सभी", "lHkh"}, {"आदेशानुसार", "vkns'kkuqlkj"}, {"विश्वासभाजन", "fo'oklHkktu"},
            {"साक्षरता", "lk{kjrk"}
        };

        public static string KrutiToUnicode(string krutiText)
        {
            if (string.IsNullOrEmpty(krutiText)) return "";

            // Protect repeated dots (e.g. ................), numeric decimals, and standalone dots using Unicode Private Use characters
            Dictionary<char, string> dotPlaceholders = new Dictionary<char, string>();
            int puaCode = 0xE000;

            string res = Regex.Replace(krutiText, @"\.{2,}", m =>
            {
                char key = (char)(puaCode++);
                dotPlaceholders[key] = m.Value;
                return key.ToString();
            });
            res = Regex.Replace(res, @"(?<=\d)\.(?=\d)", m =>
            {
                char key = (char)(puaCode++);
                dotPlaceholders[key] = ".";
                return key.ToString();
            });
            res = Regex.Replace(res, @"\.(?=\s|$)", m =>
            {
                char key = (char)(puaCode++);
                dotPlaceholders[key] = ".";
                return key.ToString();
            });

            // Protect brackets with digits/text (e.g. [1], [2], [A], [B])
            res = Regex.Replace(res, @"\[([0-9a-zA-Z\s]+)\]", m =>
            {
                char key = (char)(puaCode++);
                dotPlaceholders[key] = m.Value;
                return key.ToString();
            });

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

            res = res.Replace("अा", "आ").Replace("ाे", "ो").Replace("ाै", "ौ").Replace("ॅं", "ँ").Replace("ाॅं", "ाँ");

            // Restore protected dots and brackets
            foreach (var kvp in dotPlaceholders)
            {
                res = res.Replace(kvp.Key.ToString(), kvp.Value);
            }

            return res;
        }

        public static string UnicodeToKruti(string unicodeText)
        {
            if (string.IsNullOrEmpty(unicodeText)) return "";

            if (DirectUnicodeToKruti.ContainsKey(unicodeText))
                return DirectUnicodeToKruti[unicodeText];

            string res = unicodeText;

            // Handle खरसावाँ specifically
            res = res.Replace("खरसावाँ", "[kjlkokWa");
            res = res.Replace("सरायकेला-खरसावाँ", "ljk;dsyk&[kjlkokWa");

            // 1. Reph reordering
            res = Regex.Replace(res, @"र्((?:[\u0915-\u0939]्)*[\u0915-\u0939][\u093E-\u094C\u0901-\u0903]*)", "$1Z");

            // 2. Chhoti-i reordering
            res = Regex.Replace(res, @"((?:[\u0915-\u0939]्)*[\u0915-\u0939])ि", "f$1");

            // Direct mappings
            string[,] subs = new string[,] {
                {"औ", "vkS"}, {"ओ", "vks"}, {"आ", "vk"}, {"अ", "v"},
                {"ई", "bZ"}, {"इ", "b"}, {"ऊ", "Å"}, {"उ", "m"},
                {"ऐ", ",s"}, {"ए", ","}, {"ऋ", "_"},
                {"क्ष", "{k"}, {"त्र", "="}, {"ज्ञ", "K"}, {"श्र", "J"},
                {"क्त", "Dr"}, {"प्र", "iz"}, {"द्र", "nz"}, {"क्र", "Ø"},
                {"ट्र", "Vª"}, {"ड्र", "Mª"}, {"द्व", "}"}, {"द्य", "|"},
                {"द्ध", "Î"}, {"द्द", "Í"}, {"कृ", "Ñ"}, {"दृ", "Ò"}, {"हृ", "Ó"},
                {"रु", "®"}, {"रू", "¯"},
                {"क्", "D"}, {"ख्", "["}, {"ग्", "X"}, {"घ्", "?"},
                {"च्", "P"}, {"ज्", "T"}, {"त्", "R"}, {"थ्", "F"},
                {"ध्", "/"}, {"न्", "U"}, {"प्", "I"}, {"ब्", "C"},
                {"भ्", "H"}, {"म्", "E"}, {"ल्", "Y"}, {"व्", "O"},
                {"श्", "'"}, {"ष्", "\"k"}, {"स्", "L"}, {"ह्", "G"},
                {"क", "d"}, {"ख", "[k"}, {"ग", "x"}, {"घ", "?k"}, {"ङ", "M+"},
                {"च", "p"}, {"छ", "N"}, {"ज", "t"}, {"झ", "T"}, {"ञ", "¥"},
                {"ट", "V"}, {"ठ", "B"}, {"ड", "M"}, {"ढ", "<"}, {"ण", ".k"},
                {"त", "r"}, {"थ", "Fk"}, {"द", "n"}, {"ध", "/k"}, {"न", "u"},
                {"प", "i"}, {"फ", "Q"}, {"ब", "c"}, {"भ", "Hk"}, {"म", "e"},
                {"य", ";"}, {"र", "j"}, {"ल", "y"}, {"व", "o"},
                {"श", "'k"}, {"ष", "\"k"}, {"स", "l"}, {"ह", "g"},
                {"ा", "k"}, {"ी", "h"}, {"ु", "q"}, {"ू", "w"},
                {"ृ", "`"}, {"े", "s"}, {"ै", "S"}, {"ो", "ks"}, {"ौ", "kS"},
                {"ं", "a"}, {"ँ", "Wa"}, {"ः", "%"}, {"्", "~"}, {"्र", "z"},
                {"।", "A"}, {"०", "0"}, {"१", "1"}, {"२", "2"}, {"३", "3"},
                {"४", "4"}, {"५", "5"}, {"६", "6"}, {"७", "7"}, {"८", "8"}, {"९", "9"},
                {"-", "&"}, {"—", "&"}, {",", "]"}
            };

            for (int i = 0; i < subs.GetLength(0); i++)
            {
                res = res.Replace(subs[i, 0], subs[i, 1]);
            }

            return res;
        }
    }
}
