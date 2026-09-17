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

        // IRibbonExtensibility Implementation (Clean Professional English Toolbar)
        public string GetCustomUI(string RibbonID)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""tabKrutiDev"" label=""Kruti Dev Assistant"">
        <group id=""grpSpell"" label=""Proofing &amp; Auto-Fix"">
          <button id=""btnHighlightErrors"" label=""Highlight Errors"" size=""large"" onAction=""OnHighlightErrors"" imageMso=""HighlightColorPicker"" />
          <button id=""btnAutoFix"" label=""Auto-Fix All"" size=""large"" onAction=""OnAutoFixAll"" imageMso=""AutoCorrect"" />
          <button id=""btnHideWordSquiggles"" label=""Hide Red Lines"" size=""normal"" onAction=""OnHideEnglishSquiggles"" imageMso=""ReviewShowBalloons"" />
          <button id=""btnToggleInline"" label=""Inline Popup (On/Off)"" size=""normal"" onAction=""OnToggleInlineSuggestions"" imageMso=""GroupFont"" />
        </group>
        <group id=""grpQuickInsert"" label=""Official Letter Blocks"">
          <button id=""btnFullLetter"" label=""Full Letter Template"" size=""large"" onAction=""OnInsertFullLetter"" imageMso=""FileNewDefault"" />
          <button id=""btnSender"" label=""Sender: DSE Saraikela"" size=""normal"" onAction=""OnInsertSender"" imageMso=""MailMergeInsertAddressBlock"" />
          <button id=""btnSignature"" label=""Signature: DSE Saraikela"" size=""normal"" onAction=""OnInsertSignature"" imageMso=""SignatureLineInsert"" />
          <button id=""btnHeader"" label=""Office Letterhead"" size=""normal"" onAction=""OnInsertHeader"" imageMso=""HeaderFooterLinkToPrevious"" />
          <button id=""btnSuggestPane"" label=""Suggestions Sidebar"" size=""normal"" onAction=""OnToggleSuggestionPane"" imageMso=""Thesaurus"" />
        </group>
        <group id=""grpConvert"" label=""Font Conversion"">
          <button id=""btnToUnicode"" label=""To Unicode Hindi"" size=""normal"" onAction=""OnConvertToUnicode"" imageMso=""GroupConvert"" />
          <button id=""btnToKruti"" label=""To Kruti Dev 010"" size=""normal"" onAction=""OnConvertToKruti"" imageMso=""FontColorPicker"" />
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
                
                doc.ShowSpellingErrors = false;
                doc.ShowGrammaticalErrors = false;
                doc.Content.NoProofing = 1;

                MessageBox.Show("Word's default English red squiggly lines have been hidden.\n\nNow only actual Hindi/Kruti Dev errors will be highlighted.", "Document Cleaned", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        // 2. Highlight only real Hindi/Kruti Dev errors with DISTINCT Turquoise / Orange Marker
        public void OnHighlightErrors(IRibbonControl control)
        {
            try
            {
                if (wordApp.Documents.Count == 0)
                {
                    MessageBox.Show("Please open a document in Microsoft Word first.", "Word Connect");
                    return;
                }

                Word.Document doc = wordApp.ActiveDocument;
                
                doc.ShowSpellingErrors = false;
                doc.Content.NoProofing = 1;

                string docText = doc.Content.Text;
                var issues = spellEngine.CheckDocumentSpellingAndGrammar(docText);

                if (issues.Count == 0)
                {
                    MessageBox.Show("No spelling or grammar errors found! All words are correct.", "Check Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Clear previous highlights
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
                        range.Underline = Word.WdUnderline.wdUnderlineWavy;
                        range.Font.UnderlineColor = Word.WdColor.wdColorOrange;
                        range.HighlightColorIndex = Word.WdColorIndex.wdTurquoise;
                        highlightedCount++;
                    }
                }

                MessageBox.Show("Highlighted " + issues.Count + " actual errors in Turquoise / Orange.\n\nClick 'Auto-Fix All' to automatically fix them in your document.", "Errors Highlighted", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
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
                    MessageBox.Show("No errors found to fix.", "Auto-Fix All");
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

                doc.Content.Underline = Word.WdUnderline.wdUnderlineNone;
                doc.Content.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;

                MessageBox.Show("Successfully auto-fixed " + fixedCount + " errors in Word!", "Auto-Fix Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        public void OnToggleInlineSuggestions(IRibbonControl control)
        {
            if (inlineHook != null)
            {
                inlineHook.IsEnabled = !inlineHook.IsEnabled;
                string status = inlineHook.IsEnabled ? "ON" : "OFF";
                MessageBox.Show("Inline floating suggestion popup is now " + status + ".", "Inline Suggestions");
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
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        public void OnInsertFullLetter(IRibbonControl control)
        {
            InsertTextToWord(TemplatesData.FullOfficialLetter);
        }

        public void OnInsertSender(IRibbonControl control)
        {
            InsertTextToWord(
@"प्रेषक -
    जिला शिक्षा अधीक्षक -
    सरायकेला-खरसावाँ।");
        }

        public void OnInsertSignature(IRibbonControl control)
        {
            InsertTextToWord(
@"विश्वासभाजन

जिला शिक्षा अधीक्षक
सरायकेला-खरसावाँ।");
        }

        public void OnInsertHeader(IRibbonControl control)
        {
            InsertTextToWord(
@"कार्यालय - जिला शिक्षा अधीक्षक - सरायकेला-खरसावाँ
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
                MessageBox.Show("Converted Kruti Dev text to Unicode Hindi and copied to Clipboard!", "Conversion Complete");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
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
                MessageBox.Show("Error: " + ex.Message);
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
                MessageBox.Show("Error: " + ex.Message);
            }
        }
    }
}
