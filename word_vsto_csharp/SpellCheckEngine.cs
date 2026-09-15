using System;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace KrutiDevWordAddIn
{
    public class TrieNode
    {
        public Dictionary<char, TrieNode> Children = new Dictionary<char, TrieNode>();
        public bool IsEndOfWord = false;
        public int Frequency = 0;
        public string UnicodeWord = "";
        public string KrutiWord = "";
        public string Category = "general";
    }

    public class SpellCheckEngine
    {
        private TrieNode root = new TrieNode();
        public HashSet<string> UnicodeDictionary = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> KrutiDictionary = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private GrammarRulesEngine grammarEngine = new GrammarRulesEngine();

        public SpellCheckEngine()
        {
            LoadEmbeddedDictionary();
        }

        private void LoadEmbeddedDictionary()
        {
            // Base official words
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
                "राजनगर", "गम्हरिया", "खूँटी", "जमशेदपुर", "चाईबासा", "धनबाद", "बोकारो", "हजारीबाग"
            };

            foreach (var w in baseWords)
            {
                AddWord(w, 100);
            }
        }

        public void AddWord(string unicodeWord, int freq)
        {
            if (string.IsNullOrEmpty(unicodeWord)) return;
            string krutiWord = KrutiDevConverter.UnicodeToKruti(unicodeWord);
            
            UnicodeDictionary.Add(unicodeWord);
            if (!string.IsNullOrEmpty(krutiWord))
            {
                KrutiDictionary.Add(krutiWord);
                InsertToTrie(krutiWord, unicodeWord, freq);
            }
        }

        private void InsertToTrie(string krutiWord, string unicodeWord, int freq)
        {
            TrieNode node = root;
            foreach (char c in krutiWord)
            {
                if (!node.Children.ContainsKey(c))
                {
                    node.Children[c] = new TrieNode();
                }
                node = node.Children[c];
            }
            node.IsEndOfWord = true;
            if (freq > node.Frequency) node.Frequency = freq;
            node.UnicodeWord = unicodeWord;
            node.KrutiWord = krutiWord;
        }

        public List<string> Autocomplete(string prefixKruti, int maxResults)
        {
            List<string> results = new List<string>();
            if (string.IsNullOrEmpty(prefixKruti)) return results;

            TrieNode node = root;
            foreach (char c in prefixKruti)
            {
                if (!node.Children.ContainsKey(c)) return results;
                node = node.Children[c];
            }

            CollectWords(node, results, maxResults);
            return results;
        }

        private void CollectWords(TrieNode node, List<string> results, int maxResults)
        {
            if (results.Count >= maxResults) return;
            if (node.IsEndOfWord)
            {
                results.Add(node.UnicodeWord + " (" + node.KrutiWord + ")");
            }
            foreach (var child in node.Children.Values)
            {
                CollectWords(child, results, maxResults);
            }
        }

        public bool IsValidKrutiWord(string krutiWord)
        {
            string clean = krutiWord.Trim(new char[] { ' ', '.', ',', '%', ':', ';', '(', ')', '[', ']', '{', '}', '\'', '"', '-', 'A', ']', '[' });
            if (string.IsNullOrEmpty(clean)) return true;
            
            if (KrutiDictionary.Contains(clean)) return true;

            string uForm = KrutiDevConverter.KrutiToUnicode(clean);
            if (UnicodeDictionary.Contains(uForm)) return true;

            return false;
        }

        public string SuggestCorrection(string misspelledKruti)
        {
            string clean = misspelledKruti.Trim(new char[] { ' ', '.', ',', '%', ':', ';', '(', ')', '[', ']', '{', '}', '\'', '"', '-', 'A', ']', '[' });
            if (string.IsNullOrEmpty(clean)) return clean;

            string uMisspelled = KrutiDevConverter.KrutiToUnicode(clean);
            string bestMatch = "";
            int bestDist = int.MaxValue;

            foreach (string uWord in UnicodeDictionary)
            {
                int dist = Levenshtein(uMisspelled, uWord);
                if (dist < bestDist && dist <= 2)
                {
                    bestDist = dist;
                    bestMatch = uWord;
                }
            }

            if (!string.IsNullOrEmpty(bestMatch))
            {
                return KrutiDevConverter.UnicodeToKruti(bestMatch);
            }
            return clean;
        }

        private int Levenshtein(string s, string t)
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
