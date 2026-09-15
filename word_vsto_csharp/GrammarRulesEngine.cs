using System;
using System.Text;
using System.Collections.Generic;

namespace KrutiDevWordAddIn
{
    public class GrammarViolation
    {
        public string RuleId;
        public string Message;
        public string OriginalWordKruti;
        public string SuggestedWordKruti;
        public string OriginalWordUnicode;
        public string SuggestedWordUnicode;
        public string Context;
    }

    public class GrammarRulesEngine
    {
        private static readonly HashSet<string> VerbsPrecedingKi = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "कहा", "कहा गया", "कहा है", "कहते", "स्पष्ट है", "विदित हो", "उल्लेखनीय है",
            "सूचित किया जाता है", "आदेश दिया जाता है", "अनुरोध है", "निर्देशित किया जाता है",
            "अंकित किया", "अंकित है", "पाया गया", "देखते हुए", "प्रतीत होता है", "ज्ञात हुआ",
            "उल्लेख है", "निर्णय लिया गया", "प्रावधान है", "सहमति दी जाती है", "दिया जाता है",
            "किया जाता है", "किया गया है", "निर्देश है", "प्रार्थना है", "निवेदन है"
        };

        private static readonly HashSet<string> MeinPrefixes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "सम्बन्ध", "संबंध", "आलोक", "क्रम", "मामले", "कार्यालय", "विभाग", "जिले", "प्रखंड", "विद्यालय", "संदर्भ", "अवधि", "संचिका"
        };

        public List<GrammarViolation> CheckGrammar(string textKruti)
        {
            List<GrammarViolation> violations = new List<GrammarViolation>();
            if (string.IsNullOrEmpty(textKruti)) return violations;

            string textUnicode = KrutiDevConverter.KrutiToUnicode(textKruti);
            string[] tokens = textUnicode.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < tokens.Length; i++)
            {
                string word = tokens[i];

                // Rule 1: की vs कि (dh vs fd)
                if (word == "की")
                {
                    StringBuilder prevSb = new StringBuilder();
                    for (int j = Math.Max(0, i - 4); j < i; j++)
                    {
                        if (prevSb.Length > 0) prevSb.Append(" ");
                        prevSb.Append(tokens[j]);
                    }
                    string prevClause = prevSb.ToString();

                    foreach (var verb in VerbsPrecedingKi)
                    {
                        if (prevClause.Contains(verb))
                        {
                            violations.Add(new GrammarViolation
                            {
                                RuleId = "KI_VS_KEE_CONJUNCTION",
                                Message = "क्रिया '" + verb + "' के बाद योजक शब्द 'कि' (fd) का प्रयोग होना चाहिए, 'की' (dh) का नहीं।",
                                OriginalWordKruti = "dh",
                                SuggestedWordKruti = "fd",
                                OriginalWordUnicode = "की",
                                SuggestedWordUnicode = "कि",
                                Context = prevClause + " की ..."
                            });
                            break;
                        }
                    }
                }

                // Rule 2: में vs मैं (esa vs eSa)
                if (word == "मैं")
                {
                    string prev1 = (i > 0) ? tokens[i - 1] : "";
                    if (MeinPrefixes.Contains(prev1))
                    {
                        violations.Add(new GrammarViolation
                        {
                            RuleId = "MEIN_LOCATIVE_ERROR",
                            Message = "अधिकरण कारक के लिए 'में' (esa) का प्रयोग होना चाहिए, सर्वनाम 'मैं' (eSa) का नहीं।",
                            OriginalWordKruti = "eSa",
                            SuggestedWordKruti = "esa",
                            OriginalWordUnicode = "मैं",
                            SuggestedWordUnicode = "में",
                            Context = prev1 + " मैं ..."
                        });
                    }
                }
            }

            return violations;
        }
    }
}
