using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
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

        // Precise dictionary of verified official drafting typos -> correct spellings
        public static readonly Dictionary<string, string> KnownTypoCorrections = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // Missing reph / typo variants for 'कार्यालय'
            {"कायालय", "कार्यालय"},
            {"कायोर्लय", "कार्यालय"},
            {"कार्यलय", "कार्यालय"},
            {"कार्यालयी", "कार्यालयीय"},

            // Office letter terms
            {"पत्राक", "पत्रांक"},
            {"ज्ञापाक", "ज्ञापांक"},
            {"दिनाक", "दिनांक"},
            {"प्रतिवेधन", "प्रतिवेदन"},
            {"अनुपातन", "अनुपालन"},
            {"सचिका", "संचिका"},
            {"सञ्चिका", "संचिका"},
            {"संन्चिका", "संचिका"},
            {"टिपणी", "टिप्पणी"},
            {"टिप्पणि", "टिप्पणी"},

            // Common official drafting typos
            {"नियमैत", "नियमित"},
            {"नियमीत", "नियमित"},
            {"ममले", "मामले"},
            {"लांटती", "लौटती"},
            {"उपराक्त", "उपरोक्त"},
            {"डयता", "देयता"},
            {"सासरता", "साक्षरता"},
            {"जिंला", "जिला"},
            {"जिल्ला", "जिला"},
            {"पृष्ट", "पृष्ठ"},
            {"कारवाई", "कार्रवाई"},
            {"अदिशक", "अधीक्षक"},
            {"अधिक्षक", "अधीक्षक"},
            {"अधीछक", "अधीक्षक"},
            {"सरायकेल्ला", "सरायकेला"},
            {"सराइकेला", "सरायकेला"},
            {"खरसावां", "खरसावाँ"},
            {"खरसवां", "खरसावाँ"},
            {"झारखंड", "झारखण्ड"},
            {"राची", "राँची"},
            {"पदाधिकारीयों", "पदाधिकारियों"},
            {"प्राथमीक", "प्राथमिक"},
            {"माध्यमीक", "माध्यमिक"},
            {"शिक्षीका", "शिक्षिका"},
            {"शिक्षको", "शिक्षकों"},
            {"विदयालय", "विद्यालय"},
            {"विध्यालय", "विद्यालय"},
            {"भवदिय", "भवदीय"},
            {"हस्ताक्शर", "हस्ताक्षर"},
            {"उपस्थीती", "उपस्थिति"},
            {"उपस्थिती", "उपस्थिति"},
            {"स्वीकृती", "स्वीकृति"},
            {"अस्वीकृती", "अस्वीकृति"},
            {"निलंबण", "निलंबन"},
            {"अधोहस्ताछरी", "अधोहस्ताक्षरी"}
        };

        // Common English legal / admin terms to skip from false flagging
        private static readonly HashSet<string> EnglishKeywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "wps", "wp", "no", "nos", "slp", "cwjc", "cr", "misc", "court", "high", "state", "jharkhand",
            "ranchi", "honble", "hon'ble", "dated", "order", "page", "para", "annexure", "versus", "vs", "v/s",
            "petitioner", "respondent", "civil", "writ", "petition", "appeal", "case", "advocate", "quash",
            "disposed", "pending", "matter", "affidavit", "reply", "counter", "rejoinder", "interim", "stay", "and", "s/o", "d/o", "w/o"
        };

        public SpellCheckEngine()
        {
            LoadComprehensiveDictionary();
        }

        private void LoadComprehensiveDictionary()
        {
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

            // Comprehensive in-memory vocabulary covering all desks from 2024 data
            string[] coreWords = new string[] {
                "कार्यालय", "कार्यालय-आदेश", "शिक्षा", "विभाग", "झारखण्ड", "राँची", "सरायकेला", "खरसावाँ",
                "प्राथमिक", "माध्यमिक", "उच्च", "न्यायालय", "पदाधिकारी", "अधीक्षक", "निदेशक", "आदेश",
                "ज्ञापांक", "पत्रांक", "दिनांक", "प्रतिवेदन", "अनुपालन", "संचिका", "भवदीय", "हस्ताक्षर",
                "विषय", "प्रसंग", "महोदय", "महाशय", "अनुरोध", "आवश्यक", "कार्रवाई", "कार्यवाही", "उपरोक्त", "संलग्न",
                "वेतन", "शिक्षक", "शिक्षिका", "विद्यालय", "प्रधानाध्यापक", "प्रभारी", "छात्र", "छात्रा",
                "उपस्थिति", "अवकाश", "स्वीकृत", "अस्वीकृत", "स्पष्टीकरण", "शोकॉज", "जाँच", "निलंबन",
                "वेतनमान", "में", "मैं", "की", "कि", "के", "को", "से", "पर", "का", "है", "हैं", "था",
                "थी", "थे", "होगा", "होगी", "होंगे", "किया", "किये", "गया", "गए", "गई", "जाता", "जाती",
                "जाते", "द्वारा", "तथा", "एवं", "अथवा", "यदि", "तो", "इस", "उस", "इनके", "उनके", "सभी", "प्राप्त",
                "प्रस्तुत", "निर्देश", "निर्देशित", "समीक्षा", "बैठक", "निर्णय", "प्रस्ताव", "अनुमोदन",
                "टिप्पणी", "पंजी", "अवर", "प्रखंड", "क्षेत्रीय", "संयुक्त", "इचागढ़", "कुकडू", "चांडिल",
                "राजनगर", "गम्हरिया", "खूँटी", "जमशेदपुर", "चाईबासा", "धनबाद", "बोकारो", "हजारीबाग",
                "जिला", "पृष्ठ", "देयता", "विश्वासभाजन", "प्रेषक", "अधोहस्ताक्षरी", "चूँकि", "पूर्व",
                "बनाम", "राज्य", "सरकार", "निदेशालयीय", "स्तर", "सकारण", "खारिज", "अतएव", "अन्तरिम",
                "अंतरिम", "निदेशालय", "अग्रतर", "उचित", "प्रतीत", "होता", "अन्य", "मामले", "मांग", "माँग",
                "छायाप्रति", "सचिव", "साक्षरता", "विधि", "आलोक", "दावे", "संबंध", "सम्बन्ध", "स्थापना",
                "समिति", "नियमानुकूल", "जाए", "जाय", "दावा", "न्यायादेश", "विरुद्ध", "दायर", "हुआ", "उपर्युक्त",
                "अंकित", "करना", "याचिका", "कार्यकारी", "अंश", "निम्नवत्", "ज्ञातव्य", "मार्गदर्शन",
                "अनुलग्नक", "यथोक्त", "निकासी", "व्ययन", "नियमित", "प्रोन्नति", "पेंशन", "ग्रेच्युटी",
                "अधिसूचना", "नियमावली", "सहमति", "अनापत्ति", "शपथपत्र", "तथ्य", "विवरण", "प्रतिवादी",
                "याचिकाकर्ता", "रोक", "निस्तारण", "अवहेलना", "प्रपत्र", "प्रत्यावेदन", "अवलोकन",
                "विचारणीय", "यथोचित", "सादर", "कनीय", "लिपिक", "प्रधान", "अवगत", "विदित", "विलंब",
                "समय", "तिथि", "दिवस", "वर्ष", "माह", "सप्ताह", "दैनिक", "मासिक", "वार्षिक", "विद्यार्थी",
                "नामांकन", "छात्रवृत्ति", "भोजन", "पाठ्यपुस्तक", "परीक्षा", "मूल्यांकन", "परिणाम", "उत्तीर्ण",
                "अनुत्तीर्ण", "अंक", "प्रमाण-पत्र", "अभिलेख", "रजिस्टर", "रोकड़", "लेखा", "व्यय", "आय",
                "आवंटन", "कोषागार", "स्वीकृति", "भुगतान", "विहित", "नियम", "शर्त", "पात्रता", "योग्यता",
                "अनुभव", "प्रशिक्षण", "प्रशिक्षित", "अप्रशिक्षित", "मानदेय", "भत्ता", "महंगाई", "यात्रा",
                "आकस्मिक", "प्रतिपूरक", "प्रकाशन", "समाचार", "विज्ञप्ति", "निविदा", "कोटेशन", "आपूर्तिकर्ता",
                "अनुसार", "अनुमति", "लेने", "प्रत्युत्तर", "नहीं", "हुई", "ग्रेड", "हेतु", "प्राप्ति", "समर्पित",
                "पारित", "पोदार", "विन्देश्वरी", "स्व०", "एल०पी०ए०", "सहायक", "आचार्य", "पारा", "गैर-पारा",
                "पदस्थापन", "स्थानांतरण", "सत्यापन", "शैक्षणिक", "प्रशैक्षणिक", "सेवापुस्तिका", "सेवांत", "सेवानिवृत्ति",
                "उपदान", "मातृत्व", "उपार्जित", "स्थायीकरण", "बकाया", "पे-मैट्रिक्स", "चिकित्सा", "पुनरीक्षण", "कटौती",
                "भविष्यनिधि", "बजट", "लेखापरीक्षा", "अंकेक्षण", "आपत्ति", "पलामू", "दुमका", "देवघर", "गोड्डा", "साहिबगंज",
                "पाकुड़", "लोहरदगा", "गुमला", "सिमडेगा", "लातेहार", "गढ़वा", "चतरा", "कोडरमा", "गिरिडीह", "रामगढ़"
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

            // 1. Strict Contextual Grammar Rules Check
            CheckContextualGrammarRules(docText, issues);

            // 2. Tokenize by Whitespace & Punctuation
            HashSet<string> seenWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            string[] rawTokens = docText.Split(
                new char[] { ' ', '\t', '\r', '\n', '—', '–', '-', '(', ')', '“', '”', '‘', '’', '!', '?', ',', '।', ':', ';', '/' },
                StringSplitOptions.RemoveEmptyEntries
            );

            foreach (var raw in rawTokens)
            {
                string word = CleanToken(raw);
                if (string.IsNullOrEmpty(word) || word.Length <= 1 || seenWords.Contains(word))
                    continue;

                // Skip pure numbers or English legal headers (WPS, No, S/o)
                if (IsNumeric(word) || EnglishKeywords.Contains(word) || IsEnglishLegalHeader(word))
                    continue;

                seenWords.Add(word);

                bool isUni = ContainsDevanagari(word);

                if (isUni)
                {
                    // Unicode Devanagari Word
                    string uniWord = word;

                    // 1. Known Typo match (e.g. 'कायालय' -> 'कार्यालय', 'नियमैत' -> 'नियमित')
                    if (KnownTypoCorrections.ContainsKey(uniWord))
                    {
                        string correctedUni = KnownTypoCorrections[uniWord];
                        issues.Add(new SpellingIssue
                        {
                            OriginalWord = word,
                            SuggestedWord = correctedUni,
                            OriginalWordKruti = KrutiDevConverter.UnicodeToKruti(word),
                            SuggestedWordKruti = KrutiDevConverter.UnicodeToKruti(correctedUni),
                            Message = "अशुद्ध शब्द: '" + uniWord + "' -> सही शब्द: '" + correctedUni + "'",
                            IsUnicode = true
                        });
                        continue;
                    }

                    // 2. Missing Anusvara: 'मे' (alone, without anusvara)
                    if (uniWord == "मे")
                    {
                        issues.Add(new SpellingIssue
                        {
                            OriginalWord = word,
                            SuggestedWord = "में",
                            OriginalWordKruti = "es",
                            SuggestedWordKruti = "esa",
                            Message = "अनुस्वार छूटा: 'मे' -> 'में'",
                            IsUnicode = true
                        });
                        continue;
                    }
                }
                else
                {
                    // Kruti Dev 010 Word (ASCII typed text)
                    string uniWord = KrutiDevConverter.KrutiToUnicode(word);
                    if (string.IsNullOrEmpty(uniWord)) continue;

                    // 1. Known Typo match in Kruti Dev
                    if (KnownTypoCorrections.ContainsKey(uniWord))
                    {
                        string correctedUni = KnownTypoCorrections[uniWord];
                        string correctedKruti = KrutiDevConverter.UnicodeToKruti(correctedUni);

                        issues.Add(new SpellingIssue
                        {
                            OriginalWord = word,
                            SuggestedWord = correctedKruti,
                            OriginalWordKruti = word,
                            SuggestedWordKruti = correctedKruti,
                            Message = "अशुद्ध शब्द: '" + uniWord + "' -> सही शब्द: '" + correctedUni + "'",
                            IsUnicode = false
                        });
                        continue;
                    }

                    // 2. Missing anusvara in Kruti Dev: 'es' -> 'esa'
                    if (word == "es")
                    {
                        issues.Add(new SpellingIssue
                        {
                            OriginalWord = word,
                            SuggestedWord = "esa",
                            OriginalWordKruti = "es",
                            SuggestedWordKruti = "esa",
                            Message = "अनुस्वार छूटा: 'मे' -> 'में' (esa)",
                            IsUnicode = false
                        });
                        continue;
                    }
                }
            }

            return issues;
        }

        private void CheckContextualGrammarRules(string docText, List<SpellingIssue> issues)
        {
            // 1. संबंध मे / सम्बन्ध मे (WITHOUT anusvara) -> संबंध में / सम्बन्ध में
            MatchCollection m1 = Regex.Matches(docText, @"(?<![\u0900-\u097F])(सम्बन्ध|संबंध)\s+मे(?![\u0901-\u0903\u0900-\u097F])");
            foreach (Match m in m1)
            {
                string matchVal = m.Value;
                string replacement = matchVal.Replace("मे", "में");
                issues.Add(new SpellingIssue
                {
                    OriginalWord = matchVal,
                    SuggestedWord = replacement,
                    OriginalWordKruti = "laca/k es",
                    SuggestedWordKruti = "laca/k esa",
                    Message = "व्याकरण सुधार: '" + matchVal + "' -> '" + replacement + "'",
                    IsUnicode = true
                });
            }

            // Kruti Dev: laca/k es / lEcU/k es (WITHOUT 'a' after 'es')
            MatchCollection m1k = Regex.Matches(docText, @"\b(laca/k|lEcU/k)\s+es\b(?!a)");
            foreach (Match m in m1k)
            {
                string matchVal = m.Value;
                string replacement = matchVal + "a";
                issues.Add(new SpellingIssue
                {
                    OriginalWord = matchVal,
                    SuggestedWord = replacement,
                    OriginalWordKruti = matchVal,
                    SuggestedWordKruti = replacement,
                    Message = "व्याकरण सुधार: 'संबंध मे' -> 'संबंध में'",
                    IsUnicode = false
                });
            }

            // 2. 'है की' (after verbs) -> 'है कि'
            MatchCollection m2 = Regex.Matches(docText, @"(?<![\u0900-\u097F])(करना|किया जाता|सूचित करना|अनुरोध|विदित हो|स्पष्ट)\s+है\s+की(?![\u0900-\u097F])");
            foreach (Match m in m2)
            {
                string matchVal = m.Value;
                string replacement = matchVal.Substring(0, matchVal.Length - 2) + "कि";
                issues.Add(new SpellingIssue
                {
                    OriginalWord = matchVal,
                    SuggestedWord = replacement,
                    OriginalWordKruti = KrutiDevConverter.UnicodeToKruti(matchVal),
                    SuggestedWordKruti = KrutiDevConverter.UnicodeToKruti(replacement),
                    Message = "योजक शब्द सुधार: 'है की' -> 'है कि'",
                    IsUnicode = true
                });
            }

            // Kruti Dev: 'gS dh' (after verbs) -> 'gS fd'
            MatchCollection m2k = Regex.Matches(docText, @"\b(djuk|fd;k tkrk|lwfpr djuk|vuqjks/k)\s+gS\s+dh\b");
            foreach (Match m in m2k)
            {
                string matchVal = m.Value;
                string replacement = matchVal.Substring(0, matchVal.Length - 2) + "fd";
                issues.Add(new SpellingIssue
                {
                    OriginalWord = matchVal,
                    SuggestedWord = replacement,
                    OriginalWordKruti = matchVal,
                    SuggestedWordKruti = replacement,
                    Message = "योजक शब्द सुधार: 'है की' (gS dh) -> 'है कि' (gS fd)",
                    IsUnicode = false
                });
            }

            // 3. 'की गई था' -> 'की गई थी'
            if (docText.Contains("की गई था"))
            {
                issues.Add(new SpellingIssue
                {
                    OriginalWord = "की गई था",
                    SuggestedWord = "की गई थी",
                    OriginalWordKruti = "dh xbZ Fkk",
                    SuggestedWordKruti = "dh xbZ Fkh",
                    Message = "लिंग सहमति सुधार: 'की गई था' -> 'की गई थी'",
                    IsUnicode = true
                });
            }
            if (docText.Contains("dh xbZ Fkk"))
            {
                issues.Add(new SpellingIssue
                {
                    OriginalWord = "dh xbZ Fkk",
                    SuggestedWord = "dh xbZ Fkh",
                    OriginalWordKruti = "dh xbZ Fkk",
                    SuggestedWordKruti = "dh xbZ Fkh",
                    Message = "लिंग सहमति सुधार: 'की गई था' -> 'की गई थी'",
                    IsUnicode = false
                });
            }
        }

        private string CleanToken(string token)
        {
            if (string.IsNullOrEmpty(token)) return "";
            if (ContainsDevanagari(token))
            {
                return token.Trim(new char[] {
                    ' ', '\t', '\r', '\n', '।', ',', '.', ':', ';', '(', ')',
                    '"', '\'', '“', '”', '‘', '’', '-', '—', '–', '/', '\\', '_', '<', '>', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '०', '१', '२', '३', '४', '५', '६', '७', '८', '९'
                });
            }
            else
            {
                // Kruti Dev ASCII: only trim whitespace, standard brackets/quotes, and standalone punctuation
                return token.Trim(new char[] {
                    ' ', '\t', '\r', '\n', '(', ')', '“', '”', '‘', '’', '-', '—', '–', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9'
                });
            }
        }

        public static bool ContainsDevanagari(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            foreach (char c in text)
            {
                if (c >= '\u0900' && c <= '\u097F') return true;
            }
            return false;
        }

        private static bool IsNumeric(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            foreach (char c in text)
            {
                if (!char.IsDigit(c) && c != '०' && c != '१' && c != '२' && c != '३' && c != '४' && c != '५' && c != '६' && c != '७' && c != '८' && c != '९' && c != '.' && c != '/' && c != '-' && c != ',')
                    return false;
            }
            return true;
        }

        private static bool IsEnglishLegalHeader(string word)
        {
            if (Regex.IsMatch(word, @"^[a-zA-Z\.\(\)\/\-]+$") && (word.Contains(".") || word.Contains("/") || word.ToUpper().StartsWith("WP") || word.ToUpper().StartsWith("SLP") || word.ToUpper().StartsWith("NO") || word.ToUpper().StartsWith("S/O") || word.ToUpper().StartsWith("D/O") || word.ToUpper().StartsWith("W/O")))
                return true;
            return false;
        }

        public List<string> GetWordSuggestions(string prefix, int maxResults)
        {
            List<string> suggestions = new List<string>();
            if (string.IsNullOrEmpty(prefix)) return suggestions;

            bool isUnicode = ContainsDevanagari(prefix);
            string uniPrefix = isUnicode ? prefix : KrutiDevConverter.KrutiToUnicode(prefix);

            // 1. If prefix matches a known typo, suggest correct spelling first!
            if (KnownTypoCorrections.ContainsKey(uniPrefix))
            {
                string corrected = KnownTypoCorrections[uniPrefix];
                string kruti = KrutiDevConverter.UnicodeToKruti(corrected);
                suggestions.Add(corrected + "  [" + kruti + "]");
            }

            // 2. Exact prefix matches from dictionary
            foreach (var word in UnicodeDictionary)
            {
                if (word.StartsWith(uniPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    string candidate = word + "  [" + KrutiDevConverter.UnicodeToKruti(word) + "]";
                    if (!suggestions.Contains(candidate))
                    {
                        suggestions.Add(candidate);
                        if (suggestions.Count >= maxResults) break;
                    }
                }
            }

            return suggestions;
        }
    }
}
