using System;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace KrutiDevWordAddIn
{
    public class SpellingIssue
    {
        public string OriginalWord;
        public string SuggestedWord;
        public string OriginalWordKruti;
        public string SuggestedWordKruti;
        public string Message;
        public bool IsUnicode;
    }

    public class SpellCheckEngine
    {
        public HashSet<string> UnicodeDictionary = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> KrutiDictionary = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        
        // Common official Hindi spelling error corrections (both Unicode and Kruti)
        private static readonly Dictionary<string, string> CommonTypoCorrections = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // Frequent typos in Jharkhand office drafting
            {"जिंला", "जिला"}, {"जिल्ला", "जिला"}, {"जिलें", "जिले"},
            {"पृष्ट", "पृष्ठ"}, {"पृष्ट संख्या", "पृष्ठ संख्या"},
            {"सञ्चिका", "संचिका"}, {"संन्चिका", "संचिका"},
            {"कार्यलय", "कार्यालय"}, {"कार्यालयी", "कार्यालयीय"},
            {"कारवाई", "कार्रवाई"}, {"कार्यवाही", "कार्रवाई"},
            {"अदिशक", "अधीक्षक"}, {"अधिक्षक", "अधीक्षक"}, {"अधीछक", "अधीक्षक"},
            {"सरायकेल्ला", "सरायकेला"}, {"सराइकेला", "सरायकेला"}, {"खरसावां", "खरसावाँ"}, {"खरसवां", "खरसावाँ"},
            {"पदाधिकारीगणों", "पदाधिकारियों"}, {"पदाधिकारीयों", "पदाधिकारियों"},
            {"प्राथमीक", "प्राथमिक"}, {"माध्यमीक", "माध्यमिक"},
            {"शिक्षीका", "शिक्षिका"}, {"शिक्षको", "शिक्षकों"},
            {"विदयालय", "विद्यालय"}, {"विध्यालय", "विद्यालय"},
            {"प्रतिवेदनं", "प्रतिवेदन"}, {"अनुपालना", "अनुपालन"},
            {"शोकॉज", "स्पष्टीकरण"}, {"शोकौज", "स्पष्टीकरण"},
            {"विश्वासभाजन", "विश्वासभाजन"}, {"भवदिय", "भवदीय"},
            {"हस्ताक्शर", "हस्ताक्षर"}, {"हस्ताकक्षर", "हस्ताक्षर"},
            {"उपस्थीती", "उपस्थिति"}, {"उपस्थिती", "उपस्थिति"},
            {"स्वीकृती", "स्वीकृति"}, {"अस्वीकृती", "अस्वीकृति"},
            {"निलंबण", "निलंबन"}, {"वेतनमानं", "वेतनमान"},
            {"अधोहस्ताक्षरी", "अधोहस्ताक्षरी"}, {"अधोहस्ताछरी", "अधोहस्ताक्षरी"}
        };

        public SpellCheckEngine()
        {
            LoadEmbeddedDictionary();
        }

        private void LoadEmbeddedDictionary()
        {
            string[] baseWords = new string[] {
                "कार्यालय", "कार्यालय-आदेश", "शिक्षा", "विभाग", "झारखण्ड", "राँची", "सरायकेला", "खरसावाँ",
                "प्राथमिक", "माध्यमिक", "उच्च", "न्यायालय", "पदाधिकारी", "अधीक्षक", "निदेशक", "आदेश",
                "ज्ञापांक", "पत्रांक", "दिनांक", "प्रतिवेदन", "अनुपालन", "संचिका", "भवदीय", "हस्ताक्षर",
                "विषय", "प्रसंग", "महोदय", "अनुरोध", "आवश्यक", "कार्रवाई", "कार्यवाही", "उपरोक्त", "संलग्न",
                "वेतन", "शिक्षक", "शिक्षिका", "विद्यालय", "प्रधानाध्यापक", "प्रभारी", "छात्र", "छात्रा",
                "उपस्थिति", "अवकाश", "स्वीकृत", "अस्वीकृत", "स्पष्टीकरण", "शोकॉज", "जाँच", "निलंबन",
                "वेतनमान", "में", "मैं", "की", "कि", "के", "को", "से", "पर", "का", "है", "हैं", "था",
                "थी", "थे", "होगा", "होगी", "होंगे", "किया", "किये", "गया", "गए", "गई", "जाता", "जाती",
                "जाते", "द्वारा", "तथा", "एवं", "अथवा", "यदि", "तो", "इस", "उस", "सभी", "प्राप्त",
                "प्रस्तुत", "निर्देश", "निर्देशित", "समीक्षा", "बैठक", "निर्णय", "प्रस्ताव", "अनुमोदन",
                "टिप्पणी", "पंजी", "अवर", "प्रखंड", "क्षेत्रीय", "संयुक्त", "इचागढ़", "कुकडू", "चांडिल",
                "राजनगर", "गम्हरिया", "खूँटी", "जमशेदपुर", "चाईबासा", "धनबाद", "बोकारो", "हजारीबाग",
                "जिला", "पृष्ठ", "अधीक्षक", "विश्वासभाजन", "प्रेषक", "अधोहस्ताक्षरी"
            };

            foreach (var w in baseWords)
            {
                AddWord(w);
            }
        }

        public void AddWord(string unicodeWord)
        {
            if (string.IsNullOrEmpty(unicodeWord)) return;
            string krutiWord = KrutiDevConverter.UnicodeToKruti(unicodeWord);
            
            UnicodeDictionary.Add(unicodeWord);
            if (!string.IsNullOrEmpty(krutiWord))
            {
                KrutiDictionary.Add(krutiWord);
            }
        }

        public List<SpellingIssue> CheckDocumentSpellingAndGrammar(string docText)
        {
            List<SpellingIssue> issues = new List<SpellingIssue>();
            if (string.IsNullOrEmpty(docText)) return issues;

            // Detect if doc text is mostly Unicode or Kruti Dev
            bool isUnicode = ContainsDevanagari(docText);

            string[] tokens = docText.Split(new char[] { ' ', '\t', '\r', '\n', '।', ',', '.', ':', ';', '(', ')', '[', ']', '{', '}', '"', '\'' }, StringSplitOptions.RemoveEmptyEntries);

            HashSet<string> seenWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < tokens.Length; i++)
            {
                string word = tokens[i].Trim();
                if (string.IsNullOrEmpty(word) || word.Length <= 1 || IsNumeric(word) || seenWords.Contains(word))
                    continue;

                seenWords.Add(word);

                // 1. Check known common typos dictionary
                if (isUnicode)
                {
                    if (CommonTypoCorrections.ContainsKey(word))
                    {
                        string corrected = CommonTypoCorrections[word];
                        issues.Add(new SpellingIssue
                        {
                            OriginalWord = word,
                            SuggestedWord = corrected,
                            OriginalWordKruti = KrutiDevConverter.UnicodeToKruti(word),
                            SuggestedWordKruti = KrutiDevConverter.UnicodeToKruti(corrected),
                            Message = "अशुद्ध वर्तनी: '" + word + "' के स्थान पर '" + corrected + "' का प्रयोग करें।",
                            IsUnicode = true
                        });
                        continue;
                    }
                }
                else
                {
                    // Kruti Dev text
                    string uniWord = KrutiDevConverter.KrutiToUnicode(word);
                    if (CommonTypoCorrections.ContainsKey(uniWord))
                    {
                        string correctedUni = CommonTypoCorrections[uniWord];
                        string correctedKruti = KrutiDevConverter.UnicodeToKruti(correctedUni);
                        issues.Add(new SpellingIssue
                        {
                            OriginalWord = word,
                            SuggestedWord = correctedKruti,
                            OriginalWordKruti = word,
                            SuggestedWordKruti = correctedKruti,
                            Message = "अशुद्ध वर्तनी (कृति देव): '" + uniWord + "' (" + word + ") -> '" + correctedUni + "' (" + correctedKruti + ")",
                            IsUnicode = false
                        });
                        continue;
                    }
                }

                // 2. Fuzzy match against standard dictionary
                string targetUni = isUnicode ? word : KrutiDevConverter.KrutiToUnicode(word);
                if (!UnicodeDictionary.Contains(targetUni))
                {
                    string bestMatch = GetBestFuzzyMatch(targetUni);
                    if (!string.IsNullOrEmpty(bestMatch) && bestMatch != targetUni)
                    {
                        issues.Add(new SpellingIssue
                        {
                            OriginalWord = word,
                            SuggestedWord = isUnicode ? bestMatch : KrutiDevConverter.UnicodeToKruti(bestMatch),
                            OriginalWordKruti = isUnicode ? KrutiDevConverter.UnicodeToKruti(word) : word,
                            SuggestedWordKruti = KrutiDevConverter.UnicodeToKruti(bestMatch),
                            Message = "सुझाव: '" + targetUni + "' के लिए सही शब्द '" + bestMatch + "' हो सकता है।",
                            IsUnicode = isUnicode
                        });
                    }
                }
            }

            return issues;
        }

        private string GetBestFuzzyMatch(string uniWord)
        {
            string best = "";
            int minDistance = int.MaxValue;

            foreach (var dictWord in UnicodeDictionary)
            {
                if (Math.Abs(dictWord.Length - uniWord.Length) <= 2)
                {
                    int dist = Levenshtein(uniWord, dictWord);
                    if (dist < minDistance && dist <= 2)
                    {
                        minDistance = dist;
                        best = dictWord;
                    }
                }
            }
            return best;
        }

        public List<string> GetWordSuggestions(string prefix, int maxResults)
        {
            List<string> suggestions = new List<string>();
            if (string.IsNullOrEmpty(prefix)) return suggestions;

            bool isUnicode = ContainsDevanagari(prefix);
            string uniPrefix = isUnicode ? prefix : KrutiDevConverter.KrutiToUnicode(prefix);

            foreach (var word in UnicodeDictionary)
            {
                if (word.StartsWith(uniPrefix))
                {
                    string kruti = KrutiDevConverter.UnicodeToKruti(word);
                    suggestions.Add(word + "  [" + kruti + "]");
                    if (suggestions.Count >= maxResults) break;
                }
            }

            return suggestions;
        }

        public static bool ContainsDevanagari(string text)
        {
            foreach (char c in text)
            {
                if (c >= '\u0900' && c <= '\u097F') return true;
            }
            return false;
        }

        private static bool IsNumeric(string text)
        {
            foreach (char c in text)
            {
                if (!char.IsDigit(c) && c != '०' && c != '१' && c != '२' && c != '३' && c != '४' && c != '५' && c != '६' && c != '७' && c != '८' && c != '९')
                    return false;
            }
            return true;
        }

        private static int Levenshtein(string s, string t)
        {
            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];

            if (n == 0) return m;
            if (m == 0) return n;

            for (int i = 0; i <= n; d[i, 0] = i++) ;
            for (int j = 0; j <= m; d[0, j] = j++) ;

            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost);
                }
            }
            return d[n, m];
        }
    }
}
