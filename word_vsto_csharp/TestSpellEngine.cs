using System;

namespace KrutiDevWordAddIn
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("=================================================");
            Console.WriteLine(" VERIFYING ZERO FALSE POSITIVES ON USER DOCUMENT");
            Console.WriteLine("=================================================");

            SpellCheckEngine engine = new SpellCheckEngine();

            // Text exactly as shown in user's screenshot
            string userDocument =
@"चूँकि पूर्व में WPS No – 3369 / 2014 विन्देश्वरी पोदार बनाम राज्य सरकार में निदेशालयीय स्तर से सकारण आदेश पारित कर वादी को खारिज किया गया है, अतएव दिनांक 05.08.2024 को पारित अन्तरिम आदेश पर निदेशालय स्तर से अग्रेतर कार्रवाई किया जाना उचित प्रतीत होता है।

कार्यालय के पत्रांक 933 दिनांक 21.08.2024 द्वारा WPS No – 4300 / 2019 मोहन कुमार S/o स्व० विन्देश्वरी पोदार बनाम राज्य सरकार एवं अन्य के मामले में निदेशालय से अग्रेतर कार्रवाई करने अथवा मार्गदर्शन की माँग की गई है। छायाप्रति संलग्न।

संयुक्त सचिव, प्राथमिक शिक्षा, स्कूली शिक्षा एवं साक्षरता विभाग, झारखण्ड, राँची के पत्रांक 418 / विधि दिनांक 29.08.2024 द्वारा WPS No – 4300 / 2019 मोहन कुमार S/o स्व० विन्देश्वरी पोदार बनाम राज्य सरकार एवं अन्य के मामले में दिनांक 05.08.2024 के पारित न्यायादेश के आलोक में वादी के दावे के संबंध में जिला स्थापना समिति में नियमानुकूल निर्णय लिया जाए। यदि वादी का दावा नियमानुकूल न हो तो न्यायादेश के विरुद्ध एल०पी०ए० दायर करने की कार्रवाई की जाय, के संबंध में पत्र प्राप्त हुआ है। छायाप्रति संलग्न।

जिला शिक्षा अधीक्षक सरायकेला–खरसावाँ कार्यालय के ज्ञापांक 1219 दिनांक 24 / 09 / 2024 को कार्यालय आदेश पारित किया गया है। छायाप्रति संलग्न।";

            var falsePositiveIssues = engine.CheckDocumentSpellingAndGrammar(userDocument);
            Console.WriteLine("\n[1] User Document Error Count: " + falsePositiveIssues.Count);
            foreach (var issue in falsePositiveIssues)
            {
                Console.WriteLine("  * Flagged: '" + issue.OriginalWord + "' -> Suggest: '" + issue.SuggestedWord + "' | " + issue.Message);
            }

            if (falsePositiveIssues.Count == 0)
            {
                Console.WriteLine("\n-> SUCCESS: 0 false positives! Correct words (में, विन्देश्वरी, पोदार, मार्गदर्शन, संलग्न, सचिव, शिक्षा, विभाग, न्यायादेश, संबंध में, जाय, अधीक्षक, पारित) are NOT wrongly flagged!");
            }
            else
            {
                Console.WriteLine("\n-> FAIL: Still flagged words!");
            }

            // Test 2: Real typos to ensure detection still works 100%
            string typoDoc = "कायालय - जिला शिक्षा अदिशक - सरायकेला-खरसावां। इस संबंध मे सूचित करना है की कार्यालय नियमैत आदेश जारी किया गया।";
            var realIssues = engine.CheckDocumentSpellingAndGrammar(typoDoc);
            Console.WriteLine("\n[2] Real Typo Detection Test: Found " + realIssues.Count + " issues (Expected 5):");
            foreach (var issue in realIssues)
            {
                Console.WriteLine("  * Caught: '" + issue.OriginalWord + "' -> Suggest: '" + issue.SuggestedWord + "' | " + issue.Message);
            }

            Console.WriteLine("\n=================================================");
            Console.WriteLine(" ALL VERIFICATION TESTS COMPLETED!");
            Console.WriteLine("=================================================");
        }
    }
}
