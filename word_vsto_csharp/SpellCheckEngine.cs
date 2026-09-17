using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Reflection;

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

        // Precise dictionary of known official drafting typos -> corrections
        private static readonly Dictionary<string, string> KnownTypoCorrections = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // Drafting typos from Jharkhand education office documents
            {"कायोर्लय", "कार्यालय"}, {"कार्यलय", "कार्यालय"}, {"कार्यालयी", "कार्यालयीय"},
            {"नियमैत", "नियमित"}, {"नियमीत", "नियमित"},
            {"ममले", "मामले"}, {"मामला", "मामले"},
            {"लांटती", "लौटती"}, {"लोटती", "लौटती"},
            {"उपराक्त", "उपरोक्त"}, {"उपरोक्तक", "उपरोक्त"},
            {"डयता", "देयता"}, {"दयता", "देयता"},
            {"सासरता", "साक्षरता"}, {"साक्षारता", "साक्षरता"},
            {"जिंला", "जिला"}, {"जिल्ला", "जिला"}, {"जिलें", "जिले"},
            {"पृष्ट", "पृष्ठ"}, {"पृष्ट संख्या", "पृष्ठ संख्या"},
            {"सञ्चिका", "संचिका"}, {"संन्चिका", "संचिका"},
            {"कारवाई", "कार्रवाई"},
            {"अदिशक", "अधीक्षक"}, {"अधिक्षक", "अधीक्षक"}, {"अधीछक", "अधीक्षक"},
            {"सरायकेल्ला", "सरायकेला"}, {"सराइकेला", "सरायकेला"}, {"खरसावां", "खरसावाँ"}, {"खरसवां", "खरसावाँ"},
            {"पदाधिकारीगणों", "पदाधिकारियों"}, {"पदाधिकारीयों", "पदाधिकारियों"},
            {"प्राथमीक", "प्राथमिक"}, {"माध्यमीक", "माध्यमिक"},
            {"शिक्षीका", "शिक्षिका"}, {"शिक्षको", "शिक्षकों"},
            {"विदयालय", "विद्यालय"}, {"विध्यालय", "विद्यालय"},
            {"प्रतिवेदनं", "प्रतिवेदन"}, {"अनुपालना", "अनुपालन"},
            {"शोकौज", "स्पष्टीकरण"},
            {"भवदिय", "भवदीय"}, {"हस्ताक्शर", "हस्ताक्षर"}, {"हस्ताकक्षर", "हस्ताक्षर"},
            {"उपस्थीती", "उपस्थिति"}, {"उपस्थिती", "उपस्थिति"},
            {"स्वीकृती", "स्वीकृति"}, {"अस्वीकृती", "अस्वीकृति"},
            {"निलंबण", "निलंबन"}, {"वेतनमानं", "वेतनमान"},
            {"अधोहस्ताछरी", "अधोहस्ताक्षरी"}, {"अनुलग्न", "संलग्नक"},
            {"अनुशंसा", "अनुशंसा"}, {"अनुकुल", "अनुकूल"}
        };

        public SpellCheckEngine()
        {
            LoadComprehensiveDictionary();
        }

        private void LoadComprehensiveDictionary()
        {
            // 1. Try to load 71,000+ words from lexicon_words.txt or lexicon.json
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string txtPath = Path.Combine(baseDir, "lexicon_words.txt");
            if (!File.Exists(txtPath))
            {
                txtPath = @"d:\Lab\Project kruti dev\word_vsto_csharp\lexicon_words.txt";
            }

            if (File.Exists(txtPath))
            {
                try
                {
                    string[] lines = File.ReadAllLines(txtPath, Encoding.UTF8);
                    foreach (var line in lines)
                    {
                        string w = line.Trim();
                        if (w.Length > 0)
                        {
                            UnicodeDictionary.Add(w);
                        }
                    }
                }
                catch { }
            }

            // 2. Add base official vocabulary & Jharkhand court terms
            string[] coreWords = new string[] {
                "चूँकि", "पूर्व", "बनाम", "राज्य", "सरकार", "निदेशालयीय", "स्तर", "सकारण", "आदेश",
                "पारित", "वादी", "खारिज", "अतएव", "दिनांक", "अन्तरिम", "अंतरिम", "निदेशालय", "अग्रतर",
                "कार्रवाई", "कार्यवाही", "उचित", "प्रतीत", "होता", "कार्यालय", "पत्रांक", "ज्ञापांक",
                "अन्य", "मामले", "मांग", "छायाप्रति", "संलग्न", "संयुक्त", "सचिव", "प्राथमिक", "शिक्षा",
                "स्कूली", "साक्षरता", "विभाग", "झारखण्ड", "राँची", "विधि", "आलोक", "दावे", "संबंध", "सम्बन्ध",
                "स्थापना", "समिति", "नियमानुकूल", "निर्णय", "जाए", "यदि", "दावा", "हो", "तो", "न्यायादेश",
                "विरुद्ध", "दायरा", "दायर", "प्राप्त", "हुआ", "उपर्युक्त", "अंकित", "करना", "याचिका",
                "कार्यकारी", "अंश", "निम्नवत्", "ज्ञातव्य", "अधीक्षक", "सरायकेला", "खरसावाँ", "विश्वासभाजन",
                "भवदीय", "प्रतिवेदन", "अनुपालन", "संचिका", "टिप्पणी", "अधोहस्ताक्षरी", "स्पष्टीकरण", "शोकॉज",
                "अथवा", "मार्गदर्शन", "प्रस्ताव", "अनुमोदन", "अनुलग्नक", "यथोक्त", "निकासी", "व्ययन",
                "में", "मैं", "की", "कि", "के", "को", "से", "पर", "का", "है", "हैं", "था", "थी", "थे",
                "होगा", "होगी", "होंगे", "किया", "किये", "गया", "गए", "गई", "जाता", "जाती", "जाते",
                "द्वारा", "तथा", "एवं", "अथवा", "इस", "उस", "इनके", "उनके", "सभी", "प्रस्तुत", "निर्देश",
                "समीक्षा", "बैठक", "अवर", "प्रखंड", "क्षेत्रीय", "इचागढ़", "कुकडू", "चांडिल", "राजनगर",
                "गम्हरिया", "खूँटी", "जमशेदपुर", "चाईबासा", "धनबाद", "बोकारो", "हजारीबाग", "जिला", "पृष्ठ",
                "देयता", "वेतन", "शिक्षक", "शिक्षिका", "विद्यालय", "प्रधानाध्यापक", "प्रभारी", "छात्र",
                "छात्रा", "उपस्थिति", "अवकाश", "स्वीकृत", "अस्वीकृत", "निलंबन", "वेतनमान", "प्रोन्नति",
                "पेंशन", "ग्रेच्युटी", "अधिसूचना", "नियमावली", "सहमति", "अनापत्ति", "शपथपत्र", "तथ्य",
                "विवरण", "प्रतिवादी", "याचिकाकर्ता", "रोक", "निस्तारण", "अवहेलना", "प्रपत्र", "प्रत्यावेदन",
                "अवलोकन", "विचारणीय", "यथोचित", "सादर", "कनीय", "लिपिक", "प्रधान", "अवगत", "विदित",
                "विलंब", "समय", "तिथि", "दिवस", "वर्ष", "माह", "सप्ताह", "दैनिक", "मासिक", "वार्षिक",
                "विद्यार्थी", "नामांकन", "छात्रवृत्ति", "भोजन", "पाठ्यपुस्तक", "परीक्षा", "मूल्यांकन",
                "परिणाम", "उत्तीर्ण", "अनुत्तीर्ण", "अंक", "प्रमाण-पत्र", "अभिलेख", "पंजी", "रजिस्टर",
                "रोकड़", "लेखा", "व्यय", "आय", "आवंटन", "कोषागार", "स्वीकृति", "भुगतान", "विहित",
                "नियम", "शर्त", "पात्रता", "योग्यता", "अनुभव", "प्रशिक्षण", "प्रशिक्षित", "अप्रशिक्षित",
                "मानदेय", "भत्ता", "महंगाई", "यात्रा", "आकस्मिक", "प्रतिपूरक", "प्रकाशन", "समाचार",
                "विज्ञप्ति", "निविदा", "कोटेशन", "आपूर्तिकर्ता", "अनुसार", "अनुमति", "लेने", "प्रत्युत्तर",
                "नहीं", "हुई", "ग्रेड", "हेतु", "प्राप्ति", "समर्पित"
            };

            foreach (var w in coreWords)
            {
                UnicodeDictionary.Add(w);
                string k = KrutiDevConverter.UnicodeToKruti(w);
                if (!string.IsNullOrEmpty(k)) KrutiDictionary.Add(k);
            }
        }

        public List<SpellingIssue> CheckDocumentSpellingAndGrammar(string docText)
        {
            List<SpellingIssue> issues = new List<SpellingIssue>();
            if (string.IsNullOrEmpty(docText)) return issues;

            // 1. Check Contextual Grammar Traps first (Phrases like 'संबंध मे' -> 'संबंध में', 'की गई था' -> 'की गई थी')
            CheckContextualGrammarRules(docText, issues);

            // 2. Tokenize and Check Words
            string[] rawTokens = docText.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            HashSet<string> seenWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var raw in rawTokens)
            {
                // Clean punctuation, brackets, quotes, numbers
                string word = CleanWord(raw);
                if (string.IsNullOrEmpty(word) || word.Length <= 1 || seenWords.Contains(word))
                    continue;

                // SKIP English words, abbreviations (WPS, No, S/o, etc.) and pure numbers
                if (IsEnglishOrAlphaNumeric(word) || IsNumeric(word))
                    continue;

                seenWords.Add(word);

                bool isUnicode = ContainsDevanagari(word);
                string uniWord = isUnicode ? word : KrutiDevConverter.KrutiToUnicode(word);

                // Check 1: Known Typos Dictionary (High Accuracy)
                if (KnownTypoCorrections.ContainsKey(uniWord))
                {
                    string correctedUni = KnownTypoCorrections[uniWord];
                    string correctedKruti = KrutiDevConverter.UnicodeToKruti(correctedUni);

                    issues.Add(new SpellingIssue
                    {
                        OriginalWord = word,
                        SuggestedWord = isUnicode ? correctedUni : correctedKruti,
                        OriginalWordKruti = isUnicode ? KrutiDevConverter.UnicodeToKruti(word) : word,
                        SuggestedWordKruti = correctedKruti,
                        Message = "अशुद्ध शब्द: '" + uniWord + "' -> सही शब्द: '" + correctedUni + "'",
                        IsUnicode = isUnicode
                    });
                    continue;
                }

                // Check 2: If word is already a valid dictionary word, DO NOT FLAG IT!
                if (UnicodeDictionary.Contains(uniWord))
                {
                    continue;
                }

                // Check 3: Missing Anusvara (e.g. 'मे' -> 'में')
                if (uniWord == "मे")
                {
                    issues.Add(new SpellingIssue
                    {
                        OriginalWord = word,
                        SuggestedWord = isUnicode ? "में" : "esa",
                        OriginalWordKruti = isUnicode ? "esa" : word,
                        SuggestedWordKruti = "esa",
                        Message = "बिंदी (अनुस्वार) छूटी: 'मे' के स्थान पर 'में' (esa) का प्रयोग करें।",
                        IsUnicode = isUnicode
                    });
                    continue;
                }

                // Check 4: Fuzzy Match (Only if edit distance is 1 for words length >= 4)
                if (uniWord.Length >= 4)
                {
                    string bestMatch = GetStrictFuzzyMatch(uniWord);
                    if (!string.IsNullOrEmpty(bestMatch))
                    {
                        issues.Add(new SpellingIssue
                        {
                            OriginalWord = word,
                            SuggestedWord = isUnicode ? bestMatch : KrutiDevConverter.UnicodeToKruti(bestMatch),
                            OriginalWordKruti = isUnicode ? KrutiDevConverter.UnicodeToKruti(word) : word,
                            SuggestedWordKruti = KrutiDevConverter.UnicodeToKruti(bestMatch),
                            Message = "वर्तनी सुधार: '" + uniWord + "' -> '" + bestMatch + "'",
                            IsUnicode = isUnicode
                        });
                    }
                }
            }

            return issues;
        }

        private void CheckContextualGrammarRules(string docText, List<SpellingIssue> issues)
        {
            // 'संबंध मे' / 'सम्बन्ध मे' -> 'संबंध में'
            if (docText.Contains("संबंध मे") || docText.Contains("सम्बन्ध मे"))
            {
                issues.Add(new SpellingIssue
                {
                    OriginalWord = "संबंध मे",
                    SuggestedWord = "संबंध में",
                    OriginalWordKruti = "laca/k es",
                    SuggestedWordKruti = "laca/k esa",
                    Message = "व्याकरण सुधार: 'संबंध मे' -> 'संबंध में'",
                    IsUnicode = true
                });
            }

            // 'की गई था' -> 'की गई थी'
            if (docText.Contains("की गई था"))
            {
                issues.Add(new SpellingIssue
                {
                    OriginalWord = "की गई था",
                    SuggestedWord = "की गई थी",
                    OriginalWordKruti = "dh xbZ Fkk",
                    SuggestedWordKruti = "dh xbZ Fkh",
                    Message = "व्याकरण सुधार: 'याचिका... की गई था' -> 'की गई थी'",
                    IsUnicode = true
                });
            }

            // 'करना है की' -> 'करना है कि'
            if (docText.Contains("करना है की") || docText.Contains("किया जाता है की") || docText.Contains("सूचित करना है की"))
            {
                issues.Add(new SpellingIssue
                {
                    OriginalWord = "है की",
                    SuggestedWord = "है कि",
                    OriginalWordKruti = "gS dh",
                    SuggestedWordKruti = "gS fd",
                    Message = "योजक शब्द सुधार: 'है की' -> 'है कि'",
                    IsUnicode = true
                });
            }
        }

        private string CleanWord(string token)
        {
            if (string.IsNullOrEmpty(token)) return "";
            // Strip English/Hindi quotes, brackets, punctuation, dashes
            return token.Trim(new char[] {
                ' ', '\t', '\r', '\n', '।', ',', '.', ':', ';', '(', ')', '[', ']', '{', '}',
                '"', '\'', '“', '”', '‘', '’', '-', '—', '–', '/', '\\', '%', '_', '=', '<', '>', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '०', '१', '२', '३', '४', '५', '६', '७', '८', '९'
            });
        }

        public static bool IsEnglishOrAlphaNumeric(string text)
        {
            if (string.IsNullOrEmpty(text)) return true;
            foreach (char c in text)
            {
                if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'))
                    return true;
            }
            return false;
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
                if (!char.IsDigit(c) && c != '०' && c != '१' && c != '२' && c != '३' && c != '४' && c != '५' && c != '६' && c != '७' && c != '८' && c != '९' && c != '.' && c != '/' && c != '-')
                    return false;
            }
            return true;
        }

        private string GetStrictFuzzyMatch(string uniWord)
        {
            string best = "";
            int minDistance = int.MaxValue;

            foreach (var dictWord in UnicodeDictionary)
            {
                if (Math.Abs(dictWord.Length - uniWord.Length) <= 1)
                {
                    int dist = Levenshtein(uniWord, dictWord);
                    if (dist == 1) // Strictly 1 typo
                    {
                        return dictWord;
                    }
                }
            }
            return "";
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
