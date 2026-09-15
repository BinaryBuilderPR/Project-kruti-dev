using System;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace KrutiDevWordAddIn
{
    public class SuggestionPaneForm : Form
    {
        private Word.Application wordApp;
        private SpellCheckEngine spellEngine;
        private TextBox txtSearch;
        private ListBox lstSuggestions;
        private Label lblStatus;
        private Button btnInsertSender;
        private Button btnInsertSignature;
        private Button btnInsertHeader;
        private Button btnScanAndFix;

        public SuggestionPaneForm(Word.Application app, SpellCheckEngine engine)
        {
            this.wordApp = app;
            this.spellEngine = engine;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "शिक्षा प्रारूपक - शब्द सुझाव एवं त्वरित प्रविष्टि";
            this.Size = new Size(380, 600);
            this.StartPosition = FormStartPosition.Manual;
            this.TopMost = true;
            this.Font = new Font("Segoe UI", 9.5f);
            this.BackColor = Color.FromArgb(245, 247, 250);

            // Title Banner
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 55, BackColor = Color.FromArgb(27, 77, 62) };
            Label lblTitle = new Label
            {
                Text = "स्कूली शिक्षा एवं साक्षरता विभाग\n(सरायकेला-खरसावाँ प्रारूपक सहायक)",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
                Location = new Point(12, 8),
                AutoSize = true
            };
            topPanel.Controls.Add(lblTitle);
            this.Controls.Add(topPanel);

            // Quick Insertion Group
            GroupBox grpQuick = new GroupBox
            {
                Text = " त्वरित पत्र प्रविष्टियाँ (Quick Insert) ",
                Location = new Point(12, 65),
                Size = new Size(340, 150),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };

            btnInsertSender = new Button
            {
                Text = "✉ प्रेषक: जिला शिक्षा अधीक्षक, सरायकेला-खरसावाँ",
                Location = new Point(10, 24),
                Size = new Size(320, 34),
                BackColor = Color.FromArgb(0, 86, 179),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnInsertSender.Click += (s, e) => InsertTextToWord(
@"प्रेषक
    जिला शिक्षा अधीक्षक
    सरायकेला-खरसावाँ।");

            btnInsertSignature = new Button
            {
                Text = "✍ विश्वासभाजन: जिला शिक्षा अधीक्षक",
                Location = new Point(10, 64),
                Size = new Size(320, 34),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnInsertSignature.Click += (s, e) => InsertTextToWord(
@"विश्वासभाजन

जिला शिक्षा अधीक्षक
सरायकेला-खरसावाँ।");

            btnInsertHeader = new Button
            {
                Text = "🏢 कार्यालय शीर्ष (Office Letter Header)",
                Location = new Point(10, 104),
                Size = new Size(320, 34),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnInsertHeader.Click += (s, e) => InsertTextToWord(
@"कार्यालय - जिला शिक्षा अधीक्षक, सरायकेला-खरसावाँ
पत्रांक % ................. / दिनांक % .................");

            grpQuick.Controls.Add(btnInsertSender);
            grpQuick.Controls.Add(btnInsertSignature);
            grpQuick.Controls.Add(btnInsertHeader);
            this.Controls.Add(grpQuick);

            // Live Autocomplete Search Box
            GroupBox grpSuggest = new GroupBox
            {
                Text = " शब्द खोज एवं स्वतः पूर्ण सुझाव (IntelliSense) ",
                Location = new Point(12, 225),
                Size = new Size(340, 240),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };

            Label lblType = new Label
            {
                Text = "शब्द टाइप करें (Type prefix in Hindi or Kruti):",
                Location = new Point(10, 22),
                AutoSize = true,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Regular)
            };
            grpSuggest.Controls.Add(lblType);

            txtSearch = new TextBox
            {
                Location = new Point(10, 42),
                Size = new Size(320, 28),
                Font = new Font("Segoe UI", 11f)
            };
            txtSearch.TextChanged += TxtSearch_TextChanged;
            grpSuggest.Controls.Add(txtSearch);

            lstSuggestions = new ListBox
            {
                Location = new Point(10, 75),
                Size = new Size(320, 150),
                Font = new Font("Segoe UI", 10.5f)
            };
            lstSuggestions.DoubleClick += LstSuggestions_DoubleClick;
            grpSuggest.Controls.Add(lstSuggestions);

            this.Controls.Add(grpSuggest);

            // Bottom Scan & Fix Button
            btnScanAndFix = new Button
            {
                Text = "🔍 दस्तावेज़ में वर्तनी सुधारें (Scan & Auto-Fix Word Doc)",
                Location = new Point(12, 475),
                Size = new Size(340, 42),
                BackColor = Color.FromArgb(220, 53, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnScanAndFix.Click += BtnScanAndFix_Click;
            this.Controls.Add(btnScanAndFix);

            lblStatus = new Label
            {
                Text = "सुझाव को वर्ड में डालने के लिए सूची पर डबल-क्लिक करें।",
                Location = new Point(12, 525),
                Size = new Size(340, 30),
                ForeColor = Color.FromArgb(108, 117, 125),
                Font = new Font("Segoe UI", 8.5f, FontStyle.Italic)
            };
            this.Controls.Add(lblStatus);

            // Load initial suggestions
            UpdateSuggestions("जिला");
        }

        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            UpdateSuggestions(txtSearch.Text.Trim());
        }

        private void UpdateSuggestions(string query)
        {
            lstSuggestions.Items.Clear();
            if (string.IsNullOrEmpty(query)) return;

            var sugs = spellEngine.GetWordSuggestions(query, 12);
            foreach (var s in sugs)
            {
                lstSuggestions.Items.Add(s);
            }
        }

        private void LstSuggestions_DoubleClick(object sender, EventArgs e)
        {
            if (lstSuggestions.SelectedItem != null)
            {
                string raw = lstSuggestions.SelectedItem.ToString();
                // Extract unicode or kruti
                string word = raw.Split(new string[] { "  [" }, StringSplitOptions.None)[0].Trim();
                InsertTextToWord(word);
            }
        }

        private void BtnScanAndFix_Click(object sender, EventArgs e)
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
                    MessageBox.Show("दस्तावेज़ में कोई वर्तनी त्रुटि नहीं मिली! सभी शब्द सही हैं।", "जाँच पूर्ण", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

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

                MessageBox.Show("सफलतापूर्वक " + fixedCount + " अशुद्धियों (जैसे: 'जिंला' -> 'जिला', 'पृष्ट' -> 'पृष्ठ') को वर्ड में ठीक कर दिया गया!", "स्वतः सुधार पूर्ण", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                // Check if text is Kruti Dev or Unicode
                if (SpellCheckEngine.ContainsDevanagari(text))
                {
                    // Also provide in Kruti Dev font if current selection font is Kruti Dev
                    if (sel.Font.Name != null && sel.Font.Name.ToLower().Contains("kruti"))
                    {
                        sel.TypeText(KrutiDevConverter.UnicodeToKruti(text));
                    }
                    else
                    {
                        sel.TypeText(text);
                    }
                }
                else
                {
                    sel.Font.Name = "Kruti Dev 010";
                    sel.TypeText(text);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("त्रुटि: " + ex.Message);
            }
        }
    }
}

