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
        private GlobalInputHook inlineHook = null;
        private SuggestionPaneForm suggestionPane = null;

        // Distinct Highlight Color: Amber Orange / Cyan / Violet (distinct from default red)
        private Word.WdColor HighlightWavyColor = Word.WdColor.wdColorOrange;
        private Word.WdColorIndex HighlightBgColor = Word.WdColorIndex.wdYellow;

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            wordApp = (Word.Application)Application;
            addInInstance = AddInInst;

            try
            {
                inlineHook = new GlobalInputHook(spellEngine);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Inline hook error: " + ex.Message);
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            if (inlineHook != null)
            {
                inlineHook.Dispose();
                inlineHook = null;
            }
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
        <group id=""grpSpell"" label=""वर्तनी एवं विशिष्ट हाइलाइट"">
          <button id=""btnHighlightErrors"" label=""त्रुटियाँ विशिष्ट रंग में हाइलाइट करें"" size=""large"" onAction=""OnHighlightErrors"" imageMso=""HighlightColorPicker"" />
          <button id=""btnAutoFix"" label=""स्वतः सुधार (Auto-Fix All)"" size=""large"" onAction=""OnAutoFixAll"" imageMso=""AutoCorrect"" />
          <button id=""btnHideWordSquiggles"" label=""वर्ड की लाल लाइनें छुपाएं"" size=""normal"" onAction=""OnHideEnglishSquiggles"" imageMso=""ReviewShowBalloons"" />
          <button id=""btnToggleInline"" label=""इनलाइन पॉपअप (On/Off)"" size=""normal"" onAction=""OnToggleInlineSuggestions"" imageMso=""GroupFont"" />
        </group>
        <group id=""grpQuickInsert"" label=""सरकारी पत्र प्रविष्टियाँ"">
          <button id=""btnFullLetter"" label=""सम्पूर्ण सरकारी पत्र प्रारूप"" size=""large"" onAction=""OnInsertFullLetter"" imageMso=""FileNewDefault"" />
          <button id=""btnSender"" label=""प्रेषक: जिला शिक्षा अधीक्षक"" size=""normal"" onAction=""OnInsertSender"" imageMso=""MailMergeInsertAddressBlock"" />
          <button id=""btnSignature"" label=""विश्वासभाजन: जिला शिक्षा अधीक्षक"" size=""normal"" onAction=""OnInsertSignature"" imageMso=""SignatureLineInsert"" />
          <button id=""btnHeader"" label=""कार्यालय शीर्ष"" size=""normal"" onAction=""OnInsertHeader"" imageMso=""HeaderFooterLinkToPrevious"" />
          <button id=""btnSuggestPane"" label=""सहायक साइडबार"" size=""normal"" onAction=""OnToggleSuggestionPane"" imageMso=""Thesaurus"" />
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

        // 1. Hide default noisy English red squiggly lines across the entire Hindi document
        public void OnHideEnglishSquiggles(IRibbonControl control)
        {
            try
            {
                if (wordApp.Documents.Count == 0) return;
                Word.Document doc = wordApp.ActiveDocument;
                
                // Disable Word's English spellchecker noise on this Hindi doc
                doc.ShowSpellingErrors = false;
                doc.ShowGrammaticalErrors = false;
                doc.Content.NoProofing = 1;

                MessageBox.Show("एमएस वर्ड की अनचाही अंग्रेजी लाल लाइनें छुपा दी गई हैं!\n\nअब केवल हमारे प्रारूपक द्वारा जाँचे गए वास्तविक गलत शब्द ही अलग रंग में दिखेंगे।", "दस्तावेज़ स्वच्छ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("त्रुटि: " + ex.Message);
            }
        }

        // 2. Highlight only real Hindi/Kruti Dev errors with DISTINCT high-contrast Orange Wavy + Yellow Marker
        public void OnHighlightErrors(IRibbonControl control)
        {
            try
            {
                if (wordApp.Documents.Count == 0)
                {
                    MessageBox.Show("कृपया पहले एमएस वर्ड में कोई दस्तावेज़ खोलें।", "वर्ड कनेक्ट");
                    return;
                }

                Word.Document doc = wordApp.ActiveDocument;
                
                // First turn off Word's noisy English squiggles
                doc.ShowSpellingErrors = false;
                doc.Content.NoProofing = 1;

                string docText = doc.Content.Text;
                var issues = spellEngine.CheckDocumentSpellingAndGrammar(docText);

                if (issues.Count == 0)
                {
                    MessageBox.Show("दस्तावेज़ में कोई वर्तनी या व्याकरण त्रुटि नहीं मिली! सभी शब्द सही हैं।", "जाँच पूर्ण", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Clear previous custom highlights
                doc.Content.Underline = Word.WdUnderline.wdUnderlineNone;
                doc.Content.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;

                int highlightedCount = 0;
                foreach (var issue in issues)
                {
                    Word.Range range = doc.Content;
                    Word.Find findObj = range.Find;
                    findObj.ClearFormatting();
                    findObj.Text = issue.OriginalWord;
                    findObj.MatchCase = true;
                    findObj.Forward = true;

                    while (findObj.Execute())
                    {
                        // Use DISTINCT Cyan/Orange Wavy underline + Soft Yellow background marker
                        range.Underline = Word.WdUnderline.wdUnderlineWavy;
                        range.Font.UnderlineColor = Word.WdColor.wdColorOrange;
                        range.HighlightColorIndex = Word.WdColorIndex.wdTurquoise;
                        highlightedCount++;
                    }
                }

                MessageBox.Show("कुल " + issues.Count + " वास्तविक त्रुटियों को विशिष्ट फ़िरोज़ी (Turquoise) / नारंगी (Orange) रंग में स्पष्ट रूप से चिह्नित कर दिया गया है!\n\n(यह वर्ड के डिफ़ॉल्ट लाल रंग से अलग और स्पष्ट दिखता है)\n\nसुधारने के लिए 'स्वतः सुधार (Auto-Fix All)' बटन दबाएं।", "विशिष्ट रंग में रेखांकित", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("त्रुटि: " + ex.Message);
            }
        }

        // 3. Auto-Fix All and clear highlights
        public void OnAutoFixAll(IRibbonControl control)
        {
            try
            {
                if (wordApp.Documents.Count == 0) return;
                Word.Document doc = wordApp.ActiveDocument;
                string docText = doc.Content.Text;

                var issues = spellEngine.CheckDocumentSpellingAndGrammar(docText);
                if (issues.Count == 0)
                {
                    MessageBox.Show("कोई त्रुटि नहीं मिली।", "स्वतः सुधार");
                    return;
                }

                int fixedCount = 0;
                foreach (var issue in issues)
                {
                    Word.Find findObj = doc.Content.Find;
                    findObj.ClearFormatting();
                    findObj.Replacement.ClearFormatting();
                    findObj.Replacement.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                    findObj.Replacement.Highlight = 0;

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

                // Clear all underlines and highlights
                doc.Content.Underline = Word.WdUnderline.wdUnderlineNone;
                doc.Content.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;

                MessageBox.Show("सफलतापूर्वक " + fixedCount + " त्रुटियों को ठीक कर दिया गया!", "स्वतः सुधार पूर्ण", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("त्रुटि: " + ex.Message);
            }
        }

        public void OnToggleInlineSuggestions(IRibbonControl control)
        {
            if (inlineHook != null)
            {
                inlineHook.IsEnabled = !inlineHook.IsEnabled;
                string status = inlineHook.IsEnabled ? "चालू (ON)" : "बंद (OFF)";
                MessageBox.Show("इनलाइन सुझाव पॉपअप अब " + status + " है।", "इनलाइन स्थिति");
            }
        }

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

        public void OnInsertFullLetter(IRibbonControl control)
        {
            InsertTextToWord(TemplatesData.FullOfficialLetter);
        }

        public void OnInsertSender(IRibbonControl control)
        {
            InsertTextToWord(TemplatesData.SenderBlock);
        }

        public void OnInsertSignature(IRibbonControl control)
        {
            InsertTextToWord(TemplatesData.SignatureBlock);
        }

        public void OnInsertHeader(IRibbonControl control)
        {
            InsertTextToWord(TemplatesData.HeaderBlock);
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
                MessageBox.Show("चयनित कृति देव टेक्स्ट को यूनिकोड हिन्दी में परिवर्तित कर क्लिपबोर्ड में कॉपी कर दिया गया है!", "यूनिकोड रूपांतरण सफल");
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
