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
        private SuggestionPaneForm suggestionPane = null;

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            wordApp = (Word.Application)Application;
            addInInstance = AddInInst;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            if (suggestionPane != null && !suggestionPane.IsDisposed)
            {
                suggestionPane.Close();
                suggestionPane = null;
            }
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
      <tab id=""tabKrutiDev"" label=""शिक्षा प्रारूपक (सरायकेला-खरसावाँ)"">
        <group id=""grpSpell"" label=""वर्तनी एवं स्वतः सुधार"">
          <button id=""btnCheckDoc"" label=""दस्तावेज़ जाँचें एवं सुधारें"" size=""large"" onAction=""OnCheckDocument"" imageMso=""Spelling"" />
          <button id=""btnSuggestPane"" label=""शब्द सुझाव साइडबार"" size=""large"" onAction=""OnToggleSuggestionPane"" imageMso=""Thesaurus"" />
        </group>
        <group id=""grpQuickInsert"" label=""त्वरित पत्र प्रविष्टियाँ"">
          <button id=""btnSender"" label=""प्रेषक: जिला शिक्षा अधीक्षक"" size=""large"" onAction=""OnInsertSender"" imageMso=""MailMergeInsertAddressBlock"" />
          <button id=""btnSignature"" label=""विश्वासभाजन: जिला शिक्षा अधीक्षक"" size=""large"" onAction=""OnInsertSignature"" imageMso=""SignatureLineInsert"" />
          <button id=""btnHeader"" label=""कार्यालय शीर्ष"" size=""normal"" onAction=""OnInsertHeader"" imageMso=""HeaderFooterLinkToPrevious"" />
        </group>
        <group id=""grpConvert"" label=""फॉन्ट रूपांतरण"">
          <button id=""btnToUnicode"" label=""यूनिकोड में बदलें"" size=""normal"" onAction=""OnConvertToUnicode"" imageMso=""GroupConvert"" />
          <button id=""btnToKruti"" label=""कृति देव में बदलें"" size=""normal"" onAction=""OnConvertToKruti"" imageMso=""FontColorPicker"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        // Ribbon Callbacks
        public void OnToggleSuggestionPane(IRibbonControl control)
        {
            try
            {
                if (suggestionPane == null || suggestionPane.IsDisposed)
                {
                    suggestionPane = new SuggestionPaneForm(wordApp, spellEngine);
                }
                suggestionPane.Show();
                suggestionPane.BringToFront();
            }
            catch (Exception ex)
            {
                MessageBox.Show("त्रुटि: " + ex.Message);
            }
        }

        public void OnCheckDocument(IRibbonControl control)
        {
            try
            {
                if (wordApp.Documents.Count == 0)
                {
                    MessageBox.Show("कृपया पहले एमएस वर्ड में कोई दस्तावेज़ खोलें।", "वर्ड कनेक्ट");
                    return;
                }

                Word.Document doc = wordApp.ActiveDocument;
                string docText = doc.Content.Text;

                var issues = spellEngine.CheckDocumentSpellingAndGrammar(docText);
                if (issues.Count == 0)
                {
                    MessageBox.Show("दस्तावेज़ में कोई वर्तनी या व्याकरण त्रुटि नहीं मिली! सभी शब्द सही हैं।", "जाँच पूर्ण", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("निम्नलिखित " + issues.Count + " अशुद्धियाँ / सुझाव पाए गए:\n");
                foreach (var issue in issues)
                {
                    sb.AppendLine("• " + issue.OriginalWord + "  ->  " + issue.SuggestedWord + " (" + issue.Message + ")");
                }
                sb.AppendLine("\nक्या आप इन सभी त्रुटियों को वर्ड में स्वतः सुधारना चाहते हैं?");

                var result = MessageBox.Show(sb.ToString(), "वर्तनी एवं व्याकरण जाँच परिणाम", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    int fixedCount = 0;
                    foreach (var issue in issues)
                    {
                        Word.Find findObj = doc.Content.Find;
                        findObj.ClearFormatting();
                        findObj.Replacement.ClearFormatting();

                        object findText = issue.OriginalWord;
                        object replaceWith = issue.SuggestedWord;
                        object replaceAll = Word.WdReplace.wdReplaceAll;
                        object forward = true;
                        object matchCase = true;
                        object missing = Type.Missing;

                        findObj.Execute(
                            ref findText, ref matchCase, ref missing, ref missing, ref missing,
                            ref missing, ref forward, ref missing, ref missing, ref replaceWith,
                            ref replaceAll, ref missing, ref missing, ref missing, ref missing
                        );
                        fixedCount++;
                    }

                    MessageBox.Show("सफलतापूर्वक " + fixedCount + " त्रुटियों को वर्ड में स्वतः ठीक कर दिया गया!", "स्वतः सुधार पूर्ण", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("त्रुटि: " + ex.Message);
            }
        }

        public void OnInsertSender(IRibbonControl control)
        {
            InsertTextToWord(
@"प्रेषक,
    जिला शिक्षा अधीक्षक,
    सरायकेला-खरसावाँ।");
        }

        public void OnInsertSignature(IRibbonControl control)
        {
            InsertTextToWord(
@"विश्वासभाजन,

जिला शिक्षा अधीक्षक,
सरायकेला-खरसावाँ।");
        }

        public void OnInsertHeader(IRibbonControl control)
        {
            InsertTextToWord(
@"कार्यालय - जिला शिक्षा अधीक्षक, सरायकेला-खरसावाँ
पत्रांक % ................. / दिनांक % .................");
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

        private void InsertTextToWord(string text)
        {
            try
            {
                if (wordApp.Documents.Count == 0)
                {
                    wordApp.Documents.Add();
                }
                Word.Selection sel = wordApp.Selection;
                if (sel.Font.Name != null && sel.Font.Name.ToLower().Contains("kruti"))
                {
                    sel.TypeText(KrutiDevConverter.UnicodeToKruti(text) + "\n");
                }
                else
                {
                    sel.TypeText(text + "\n");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("त्रुटि: " + ex.Message);
            }
        }
    }
}
