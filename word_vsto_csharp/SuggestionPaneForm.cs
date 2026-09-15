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
        private Button btnFullLetter;
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
            this.Text = "शिक्षा प्रारूपक - सरायकेला-खरसावाँ";
            this.Size = new Size(390, 640);
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
                Text = " त्वरित सरकारी पत्र प्रविष्टियाँ (1-Click) ",
                Location = new Point(12, 65),
                Size = new Size(350, 180),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };

            btnFullLetter = new Button
            {
                Text = "📄 सम्पूर्ण सरकारी पत्र प्रारूप (Full Letter Template)",
                Location = new Point(10, 22),
                Size = new Size(330, 34),
                BackColor = Color.FromArgb(13, 110, 253),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnFullLetter.Click += (s, e) => InsertTextToWord(TemplatesData.FullOfficialLetter);

            btnInsertSender = new Button
            {
                Text = "✉ प्रेषक: जिला शिक्षा अधीक्षक, सरायकेला-खरसावाँ",
                Location = new Point(10, 60),
                Size = new Size(330, 34),
                BackColor = Color.FromArgb(0, 86, 179),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnInsertSender.Click += (s, e) => InsertTextToWord(TemplatesData.SenderBlock);

            btnInsertSignature = new Button
            {
                Text = "✍ विश्वासभाजन: जिला शिक्षा अधीक्षक, सरायकेला-खरसावाँ",
                Location = new Point(10, 98),
                Size = new Size(330, 34),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnInsertSignature.Click += (s, e) => InsertTextToWord(TemplatesData.SignatureBlock);

            btnInsertHeader = new Button
            {
                Text = "🏢 कार्यालय शीर्ष: सरायकेला-खरसावाँ",
                Location = new Point(10, 136),
                Size = new Size(330, 34),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnInsertHeader.Click += (s, e) => InsertTextToWord(TemplatesData.HeaderBlock);

            grpQuick.Controls.Add(btnFullLetter);
            grpQuick.Controls.Add(btnInsertSender);
            grpQuick.Controls.Add(btnInsertSignature);
            grpQuick.Controls.Add(btnInsertHeader);
            this.Controls.Add(grpQuick);

            // Live Autocomplete Search Box
            GroupBox grpSuggest = new GroupBox
            {
                Text = " शब्द खोज एवं स्वतः पूर्ण सुझाव (IntelliSense) ",
                Location = new Point(12, 255),
                Size = new Size(350, 235),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };

            Label lblType = new Label
            {
                Text = "शब्द टाइप करें (Type in Hindi or Kruti):",
                Location = new Point(10, 22),
                AutoSize = true,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Regular)
            };
            grpSuggest.Controls.Add(lblType);

            txtSearch = new TextBox
            {
                Location = new Point(10, 42),
                Size = new Size(330, 28),
                Font = new Font("Segoe UI", 11f)
            };
            txtSearch.TextChanged += TxtSearch_TextChanged;
            grpSuggest.Controls.Add(txtSearch);

            lstSuggestions = new ListBox
            {
                Location = new Point(10, 75),
                Size = new Size(330, 150),
                Font = new Font("Segoe UI", 10.5f)
            };
            lstSuggestions.DoubleClick += LstSuggestions_DoubleClick;
            grpSuggest.Controls.Add(lstSuggestions);

            this.Controls.Add(grpSuggest);

            // Bottom Scan & Fix Button
            btnScanAndFix = new Button
            {
                Text = "🔍 दस्तावेज़ में वर्तनी सुधारें (Scan & Auto-Fix Doc)",
                Location = new Point(12, 498),
                Size = new Size(350, 42),
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
                Location = new Point(12, 545),
                Size = new Size(350, 25),
                ForeColor = Color.FromArgb(108, 117, 125),
                Font = new Font("Segoe UI", 8.5f, FontStyle.Italic)
            };
            this.Controls.Add(lblStatus);

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

                MessageBox.Show("सफलतापूर्वक " + fixedCount + " अशुद्धियों को वर्ड में ठीक कर दिया गया!", "स्वतः सुधार पूर्ण", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
