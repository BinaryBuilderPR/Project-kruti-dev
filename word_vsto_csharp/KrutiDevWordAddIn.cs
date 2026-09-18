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

            try
            {
                wordApp.WindowBeforeRightClick += WordApp_WindowBeforeRightClick;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Right click hook error: " + ex.Message);
            }
        }

        // Right-Click Context Menu Quick-Fix Provider (On-demand right click)
        private void WordApp_WindowBeforeRightClick(Word.Selection sel, ref bool cancel)
        {
            try
            {
                if (sel == null || sel.Words.Count == 0) return;
                Word.Range wordRange = sel.Words[1];
                string rawText = wordRange.Text;
                if (string.IsNullOrEmpty(rawText)) return;

                string selWord = rawText.Trim(new char[] { ' ', '\t', '\r', '\n', '।', ',', '.', ':', ';', '(', ')', '"', '\'', '-', '—', '–' });
                if (string.IsNullOrEmpty(selWord)) return;

                bool isUni = SpellCheckEngine.ContainsDevanagari(selWord);
                string uniWord = isUni ? selWord : KrutiDevConverter.KrutiToUnicode(selWord);

                string correctionUni = "";
                if (SpellCheckEngine.KnownTypoCorrections.ContainsKey(uniWord))
                {
                    correctionUni = SpellCheckEngine.KnownTypoCorrections[uniWord];
                }
                else if (uniWord == "मे" || selWord == "es")
                {
                    correctionUni = "में";
                }

                if (!string.IsNullOrEmpty(correctionUni))
                {
                    string correctionText = isUni ? correctionUni : KrutiDevConverter.UnicodeToKruti(correctionUni);

                    Microsoft.Office.Core.CommandBar contextMenu = wordApp.CommandBars["Text"];
                    if (contextMenu != null)
                    {
                        for (int i = contextMenu.Controls.Count; i >= 1; i--)
                        {
                            try
                            {
                                if (contextMenu.Controls[i].Tag == "KrutiDevQuickFix")
                                {
                                    contextMenu.Controls[i].Delete(Type.Missing);
                                }
                            }
                            catch { }
                        }

                        Microsoft.Office.Core.CommandBarButton btnFix = (Microsoft.Office.Core.CommandBarButton)contextMenu.Controls.Add(
                            Microsoft.Office.Core.MsoControlType.msoControlButton, Type.Missing, Type.Missing, 1, true);
                        btnFix.Caption = "✨ Fix to: " + correctionUni + " (" + correctionText + ")";
                        btnFix.Tag = "KrutiDevQuickFix";
                        btnFix.BeginGroup = true;

                        btnFix.Click += (Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault) =>
                        {
                            try
                            {
                                wordRange.Text = correctionText + " ";
                                wordRange.Underline = Word.WdUnderline.wdUnderlineNone;
                                wordRange.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
                            }
                            catch { }
                        };
                    }
                }
            }
            catch { }
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

        // IRibbonExtensibility Implementation (Clean Professional English Ribbon Toolbar)
        public string GetCustomUI(string RibbonID)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""tabKrutiDev"" label=""Kruti Dev Assistant"">
        <group id=""grpSpell"" label=""Proofing &amp; Auto-Fix"">
          <button id=""btnAutoFix"" label=""Auto-Fix All"" size=""large"" onAction=""OnAutoFixAll"" imageMso=""AutoCorrect"" />
          <button id=""btnHighlightErrors"" label=""Highlight Errors"" size=""large"" onAction=""OnHighlightErrors"" imageMso=""HighlightColorPicker"" />
          <button id=""btnClearHighlights"" label=""Clear Highlights"" size=""normal"" onAction=""OnClearHighlights"" imageMso=""ClearFormatting"" />
          <button id=""btnHideWordSquiggles"" label=""Hide Red Lines"" size=""normal"" onAction=""OnHideEnglishSquiggles"" imageMso=""ReviewShowBalloons"" />
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

        public void OnClearHighlights(IRibbonControl control)
        {
            try
            {
                if (wordApp.Documents.Count == 0) return;
                Word.Document doc = wordApp.ActiveDocument;
                
                doc.Content.Underline = Word.WdUnderline.wdUnderlineNone;
                doc.Content.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;

                MessageBox.Show("All error highlights and underlines have been cleared.", "Highlights Cleared", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Kruti Dev Assistant");
            }
        }

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
                MessageBox.Show("Error: " + ex.Message, "Kruti Dev Assistant");
            }
        }

        public void OnHighlightErrors(IRibbonControl control)
        {
            try
            {
                if (wordApp.Documents.Count == 0) return;
                Word.Document doc = wordApp.ActiveDocument;
                string docText = doc.Content.Text;

                // Clear existing highlights first
                doc.Content.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
                doc.Content.Underline = Word.WdUnderline.wdUnderlineNone;

                var issues = spellEngine.CheckDocumentSpellingAndGrammar(docText);

                if (issues.Count == 0)
                {
                    MessageBox.Show("No spelling or grammar errors found! All words are correct.", "Check Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int highlightedCount = 0;
                foreach (var issue in issues)
                {
                    Word.Range range = doc.Content;
                    Word.Find findObj = range.Find;
                    findObj.ClearFormatting();
                    findObj.Text = issue.OriginalWord;
                    findObj.MatchCase = true;
                    findObj.MatchWholeWord = false;
                    findObj.Forward = true;
                    findObj.Wrap = Word.WdFindWrap.wdFindStop;

                    while (findObj.Execute())
                    {
                        range.Underline = Word.WdUnderline.wdUnderlineWavy;
                        range.Font.UnderlineColor = Word.WdColor.wdColorOrange;
                        range.HighlightColorIndex = Word.WdColorIndex.wdTurquoise;
                        highlightedCount++;
                        range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }

                MessageBox.Show("Found and highlighted " + issues.Count + " error pattern(s) (" + highlightedCount + " instance(s)) in Turquoise / Orange.\n\nClick 'Auto-Fix All' to automatically correct them and remove underlines.", "Errors Highlighted", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Kruti Dev Assistant");
            }
        }

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
                    doc.Content.Underline = Word.WdUnderline.wdUnderlineNone;
                    doc.Content.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
                    MessageBox.Show("No errors found to fix. All underlines cleared.", "Auto-Fix All", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    object wrap = Word.WdFindWrap.wdFindContinue;
                    object missing = Type.Missing;

                    findObj.Execute(
                        ref findText, ref matchCase, ref missing, ref missing, ref missing,
                        ref missing, ref forward, ref wrap, ref missing, ref replaceWith,
                        ref replaceAll, ref missing, ref missing, ref missing, ref missing
                    );
                    fixedCount++;
                }

                doc.Content.Underline = Word.WdUnderline.wdUnderlineNone;
                doc.Content.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;

                MessageBox.Show("Successfully auto-fixed " + fixedCount + " error pattern(s) in Word!\n\nAll underlines and highlights have been automatically removed.", "Auto-Fix Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Kruti Dev Assistant");
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
                MessageBox.Show("Error: " + ex.Message, "Kruti Dev Assistant");
            }
        }

        public void OnInsertFullLetter(IRibbonControl control)
        {
            InsertTextToWord(TemplatesData.FullOfficialLetter);
        }

        public void OnInsertSender(IRibbonControl control)
        {
            InsertTextToWord("izs\"kd]\n    funs'kd ¼izk0f'k0½\n    Ldwyh f'k{kk ,oa lk{kjrk foHkkx]\n    >kj[k.M] jk¡phA");
        }

        public void OnInsertSignature(IRibbonControl control)
        {
            InsertTextToWord("fo'oklHkktu\n\n\nftyk f'k{kk v/kh{kd\nljk;dsyk&[kjlkokaA");
        }

        public void OnInsertHeader(IRibbonControl control)
        {
            InsertTextToWord("dk;kZy; ftyk f'k{kk v/kh{kd] ljk;dsyk&[kjlkoka\n¼Ldwyh f'k{kk ,oa lk{kjrk foHkkx] >kj[k.M ljdkj½");
        }

        public void OnConvertToUnicode(IRibbonControl control)
        {
            try
            {
                if (wordApp.Documents.Count == 0) return;
                Word.Selection sel = wordApp.Selection;
                if (sel == null || string.IsNullOrEmpty(sel.Text))
                {
                    MessageBox.Show("Please select some Kruti Dev text in your document first.", "No Text Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string converted = KrutiDevConverter.KrutiToUnicode(sel.Text);
                sel.Text = converted;
                sel.Font.Name = "Mangal";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error converting to Unicode: " + ex.Message, "Kruti Dev Assistant");
            }
        }

        public void OnConvertToKruti(IRibbonControl control)
        {
            try
            {
                if (wordApp.Documents.Count == 0) return;
                Word.Selection sel = wordApp.Selection;
                if (sel == null || string.IsNullOrEmpty(sel.Text))
                {
                    MessageBox.Show("Please select some Unicode Hindi text in your document first.", "No Text Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string converted = KrutiDevConverter.UnicodeToKruti(sel.Text);
                sel.Text = converted;
                sel.Font.Name = "Kruti Dev 010";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error converting to Kruti Dev: " + ex.Message, "Kruti Dev Assistant");
            }
        }

        private void InsertTextToWord(string text)
        {
            try
            {
                if (wordApp.Documents.Count == 0) return;
                Word.Selection sel = wordApp.Selection;
                if (sel != null)
                {
                    sel.Text = text + "\n";
                    sel.Font.Name = "Kruti Dev 010";
                    sel.Font.Size = 14;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error inserting template: " + ex.Message, "Kruti Dev Assistant");
            }
        }
    }
}
