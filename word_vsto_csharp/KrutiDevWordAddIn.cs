using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace KrutiDevWordAddIn
{
    [Guid("A789B1C2-3D4E-5F6A-7B8C-9D0E1F2A3B4C")]
    [ProgId("KrutiDevWordAddIn.Connect")]
    [ComVisible(true)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private Word.Application wordApp;
        private object addInInstance;
        private SpellCheckEngine spellEngine = new SpellCheckEngine();
        private GrammarRulesEngine grammarEngine = new GrammarRulesEngine();

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            wordApp = (Word.Application)Application;
            addInInstance = AddInInst;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            wordApp = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // IRibbonExtensibility Implementation
        public string GetCustomUI(string RibbonID)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""tabKrutiDev"" label=""शिक्षा प्रारूपक (कृति देव)"">
        <group id=""grpSpell"" label=""वर्तनी एवं व्याकरण"">
          <button id=""btnCheckDoc"" label=""दस्तावेज़ जाँचें"" size=""large"" onAction=""OnCheckDocument"" imageMso=""Spelling"" />
          <button id=""btnAutoFix"" label=""स्वतः सुधार (की/कि, में/मैं)"" size=""large"" onAction=""OnAutoFixGrammar"" imageMso=""AutoCorrect"" />
          <button id=""btnSuggest"" label=""सुझाव देखें"" size=""normal"" onAction=""OnSuggestWord"" imageMso=""Thesaurus"" />
        </group>
        <group id=""grpConvert"" label=""कृति देव ⇄ यूनिकोड"">
          <button id=""btnToUnicode"" label=""यूनिकोड में बदलें"" size=""large"" onAction=""OnConvertToUnicode"" imageMso=""GroupConvert"" />
          <button id=""btnToKruti"" label=""कृति देव 010 में बदलें"" size=""large"" onAction=""OnConvertToKruti"" imageMso=""FontColorPicker"" />
        </group>
        <group id=""grpTemplates"" label=""सरकारी प्रारूप (1-Click)"">
          <button id=""btnTplOrder"" label=""कार्यालय-आदेश"" size=""normal"" onAction=""OnInsertOrder"" imageMso=""FileNewDefault"" />
          <button id=""btnTplNotice"" label=""कारण पृच्छा (Show-Cause)"" size=""normal"" onAction=""OnInsertNotice"" imageMso=""AlertSettings"" />
          <button id=""btnTplSOF"" label=""माननीय न्यायालय तथ्य-विवरण (SOF)"" size=""normal"" onAction=""OnInsertSOF"" imageMso=""ReviewAcceptChange"" />
          <button id=""btnTplNoting"" label=""संचिका टिप्पणी"" size=""normal"" onAction=""OnInsertNoting"" imageMso=""NewTask"" />
          <button id=""btnTplATR"" label=""अनुपालन प्रतिवेदन"" size=""normal"" onAction=""OnInsertATR"" imageMso=""SendCopy"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        // Ribbon Callbacks
        public void OnCheckDocument(IRibbonControl control)
        {
            try
            {
                if (wordApp.Documents.Count == 0) return;
                Word.Document doc = wordApp.ActiveDocument;
                string text = doc.Content.Text;

                var grammarIssues = grammarEngine.CheckGrammar(text);
                int count = grammarIssues.Count;

                if (count == 0)
                {
                    MessageBox.Show("दस्तावेज़ में व्याकरण की कोई त्रुटि (की/कि, में/मैं) नहीं मिली!", "स्कैन पूर्ण", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("निम्नलिखित " + count + " व्याकरण संबंधी सुझाव मिले:\n");
                    foreach (var issue in grammarIssues)
                    {
                        sb.AppendLine("• " + issue.Message + " (संदर्भ: " + issue.Context + ")");
                    }
                    sb.AppendLine("\nक्या आप इन त्रुटियों को स्वतः सुधारना चाहते हैं?");
                    
                    var result = MessageBox.Show(sb.ToString(), "व्याकरण जाँच परिणाम", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        OnAutoFixGrammar(control);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("त्रुटि: " + ex.Message);
            }
        }

        public void OnAutoFixGrammar(IRibbonControl control)
        {
            try
            {
                if (wordApp.Documents.Count == 0) return;
                Word.Document doc = wordApp.ActiveDocument;
                string text = doc.Content.Text;

                var grammarIssues = grammarEngine.CheckGrammar(text);
                int fixedCount = 0;

                foreach (var issue in grammarIssues)
                {
                    Word.Find findObj = doc.Content.Find;
                    findObj.ClearFormatting();
                    findObj.Replacement.ClearFormatting();
                    findObj.Replacement.Font.Name = "Kruti Dev 010";

                    object findText = issue.OriginalWordKruti;
                    object replaceWith = issue.SuggestedWordKruti;
                    object replaceOne = Word.WdReplace.wdReplaceOne;
                    object forward = true;
                    object matchCase = true;
                    object missing = Type.Missing;

                    findObj.Execute(
                        ref findText, ref matchCase, ref missing, ref missing, ref missing,
                        ref missing, ref forward, ref missing, ref missing, ref replaceWith,
                        ref replaceOne, ref missing, ref missing, ref missing, ref missing
                    );
                    fixedCount++;
                }

                MessageBox.Show("सफलतापूर्वक " + fixedCount + " त्रुटियों का स्वतः सुधार कर दिया गया!", "स्वतः सुधार पूर्ण", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("त्रुटि: " + ex.Message);
            }
        }

        public void OnSuggestWord(IRibbonControl control)
        {
            try
            {
                if (wordApp.Documents.Count == 0) return;
                string selectedText = wordApp.Selection.Text.Trim();
                if (string.IsNullOrEmpty(selectedText))
                {
                    MessageBox.Show("कृपया पहले वर्ड में किसी शब्द का चयन करें।", "चयन आवश्यक");
                    return;
                }

                string suggestionKruti = spellEngine.SuggestCorrection(selectedText);
                string suggestionUni = KrutiDevConverter.KrutiToUnicode(suggestionKruti);

                var res = MessageBox.Show("चयनित शब्द: " + selectedText + " (" + KrutiDevConverter.KrutiToUnicode(selectedText) + ")\n"
                    + "सुझाया गया सही शब्द: " + suggestionUni + " (" + suggestionKruti + ")\n\n"
                    + "क्या आप इसे वर्ड में प्रतिस्थापित (Replace) करना चाहते हैं?",
                    "वर्तनी सुधार सुझाव", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (res == DialogResult.Yes)
                {
                    wordApp.Selection.Font.Name = "Kruti Dev 010";
                    wordApp.Selection.TypeText(suggestionKruti);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("त्रुटि: " + ex.Message);
            }
        }

        public void OnConvertToUnicode(IRibbonControl control)
        {
            try
            {
                if (wordApp.Documents.Count == 0) return;
                string selText = wordApp.Selection.Text;
                if (string.IsNullOrEmpty(selText.Trim()))
                {
                    selText = wordApp.ActiveDocument.Content.Text;
                }

                string unicodeText = KrutiDevConverter.KrutiToUnicode(selText);
                Clipboard.SetText(unicodeText);
                MessageBox.Show("चयनित कृति देव टेक्स्ट को यूनिकोड हिन्दी में परिवर्तित कर क्लिपबोर्ड में कॉपी कर दिया गया है!\n\n(आप इसे सीधे ई-कल्याण/एचआरएमएस या ईमेल में पेस्ट कर सकते हैं)", "यूनिकोड रूपांतरण सफल");
            }
            catch (Exception ex)
            {
                MessageBox.Show("त्रुटि: " + ex.Message);
            }
        }

        public void OnConvertToKruti(IRibbonControl control)
        {
            try
            {
                if (wordApp.Documents.Count == 0) return;
                string selText = wordApp.Selection.Text;
                if (string.IsNullOrEmpty(selText.Trim())) return;

                string krutiText = KrutiDevConverter.UnicodeToKruti(selText);
                wordApp.Selection.Font.Name = "Kruti Dev 010";
                wordApp.Selection.TypeText(krutiText);
            }
            catch (Exception ex)
            {
                MessageBox.Show("त्रुटि: " + ex.Message);
            }
        }

        public void OnInsertOrder(IRibbonControl control)
        {
            InsertTemplate(TemplatesData.OfficeOrder);
        }

        public void OnInsertNotice(IRibbonControl control)
        {
            InsertTemplate(TemplatesData.ShowCauseNotice);
        }

        public void OnInsertSOF(IRibbonControl control)
        {
            InsertTemplate(TemplatesData.StatementOfFacts);
        }

        public void OnInsertNoting(IRibbonControl control)
        {
            InsertTemplate(TemplatesData.FileNoting);
        }

        public void OnInsertATR(IRibbonControl control)
        {
            InsertTemplate(TemplatesData.ComplianceReport);
        }

        private void InsertTemplate(string templateContentKruti)
        {
            try
            {
                if (wordApp.Documents.Count == 0)
                {
                    wordApp.Documents.Add();
                }
                Word.Selection sel = wordApp.Selection;
                sel.Font.Name = "Kruti Dev 010";
                sel.Font.Size = 13;
                sel.TypeText(templateContentKruti);
            }
            catch (Exception ex)
            {
                MessageBox.Show("त्रुटि: " + ex.Message);
            }
        }
    }
}
