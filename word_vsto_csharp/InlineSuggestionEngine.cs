using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace KrutiDevWordAddIn
{
    /// <summary>
    /// Google Input Tools / VS Code style Inline Floating Suggestion Window.
    /// Floats directly under the typing cursor in Microsoft Word and all Windows apps.
    /// </summary>
    public class InlineSuggestionForm : Form
    {
        private ListBox lstSuggestions;
        private SpellCheckEngine spellEngine;
        private string currentBuffer = "";
        private List<string> currentCandidates = new List<string>();
        private List<string> currentCandidatesKruti = new List<string>();

        // Win32 Caret and Hook APIs
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int WS_EX_TOPMOST = 0x00000008;
        private const int WS_EX_TOOLWINDOW = 0x00000080;

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= WS_EX_NOACTIVATE | WS_EX_TOPMOST | WS_EX_TOOLWINDOW;
                return cp;
            }
        }

        protected override bool ShowWithoutActivation
        {
            get { return true; }
        }

        public InlineSuggestionForm(SpellCheckEngine engine)
        {
            this.spellEngine = engine;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.Manual;
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.BackColor = Color.FromArgb(220, 224, 230);
            this.Size = new Size(270, 165);
            this.Padding = new Padding(1);
            this.DoubleBuffered = true;

            Panel borderPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(2)
            };

            lstSuggestions = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 11f, FontStyle.Regular),
                ItemHeight = 28,
                DrawMode = DrawMode.OwnerDrawFixed,
                BackColor = Color.White
            };
            lstSuggestions.DrawItem += LstSuggestions_DrawItem;
            borderPanel.Controls.Add(lstSuggestions);
            this.Controls.Add(borderPanel);
        }

        private void LstSuggestions_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= currentCandidates.Count) return;

            bool isSelected = (e.Index == 0); // Top suggestion highlighted like Google Input Tools
            Color bg = isSelected ? Color.FromArgb(215, 238, 255) : Color.White;
            Color fg = Color.FromArgb(20, 20, 20);
            Color numColor = Color.FromArgb(90, 90, 90);
            Color krutiColor = Color.FromArgb(120, 120, 120);

            using (SolidBrush bgBrush = new SolidBrush(bg))
            {
                e.Graphics.FillRectangle(bgBrush, e.Bounds);
            }

            if (isSelected)
            {
                using (Pen borderPen = new Pen(Color.FromArgb(140, 195, 255)))
                {
                    e.Graphics.DrawRectangle(borderPen, e.Bounds.X, e.Bounds.Y, e.Bounds.Width - 1, e.Bounds.Height - 1);
                }
            }

            // Draw number (1..5)
            string numStr = (e.Index + 1).ToString() + ".";
            using (SolidBrush numBrush = new SolidBrush(numColor))
            {
                e.Graphics.DrawString(numStr, new Font("Segoe UI", 9.5f, FontStyle.Bold), numBrush, e.Bounds.X + 4, e.Bounds.Y + 5);
            }

            // Draw Hindi Word
            string wordStr = currentCandidates[e.Index];
            using (SolidBrush textBrush = new SolidBrush(fg))
            {
                e.Graphics.DrawString(wordStr, new Font("Segoe UI", 11f, FontStyle.Bold), textBrush, e.Bounds.X + 26, e.Bounds.Y + 3);
            }

            // Draw Kruti Remington Hint (gray)
            if (e.Index < currentCandidatesKruti.Count && !string.IsNullOrEmpty(currentCandidatesKruti[e.Index]))
            {
                string kHint = "[" + currentCandidatesKruti[e.Index] + "]";
                using (SolidBrush hintBrush = new SolidBrush(krutiColor))
                {
                    SizeF wordSize = e.Graphics.MeasureString(wordStr, new Font("Segoe UI", 11f, FontStyle.Bold));
                    e.Graphics.DrawString(kHint, new Font("Segoe UI", 8.5f, FontStyle.Italic), hintBrush, e.Bounds.X + 28 + (int)wordSize.Width + 4, e.Bounds.Y + 6);
                }
            }
        }

        public void UpdateBufferAndSuggestions(string buffer, Point caretPos)
        {
            this.currentBuffer = buffer;

            if (string.IsNullOrEmpty(buffer) || buffer.Trim().Length == 0)
            {
                this.Hide();
                return;
            }

            var sugs = spellEngine.GetWordSuggestions(buffer, 5);
            if (sugs.Count == 0)
            {
                // Fallback direct conversion
                string direct = SpellCheckEngine.ContainsDevanagari(buffer) ? buffer : KrutiDevConverter.KrutiToUnicode(buffer);
                if (!string.IsNullOrEmpty(direct))
                {
                    sugs.Add(direct + "  [" + KrutiDevConverter.UnicodeToKruti(direct) + "]");
                }
            }

            if (sugs.Count == 0)
            {
                this.Hide();
                return;
            }

            currentCandidates.Clear();
            currentCandidatesKruti.Clear();

            foreach (var s in sugs)
            {
                string[] parts = s.Split(new string[] { "  [" }, StringSplitOptions.None);
                string uWord = parts[0].Trim();
                string kWord = parts.Length > 1 ? parts[1].Replace("]", "").Trim() : KrutiDevConverter.UnicodeToKruti(uWord);
                currentCandidates.Add(uWord);
                currentCandidatesKruti.Add(kWord);
            }

            lstSuggestions.Items.Clear();
            foreach (var c in currentCandidates)
            {
                lstSuggestions.Items.Add(c);
            }

            int itemHeight = 28;
            int totalHeight = Math.Max(35, currentCandidates.Count * itemHeight + 6);
            this.Size = new Size(270, totalHeight);

            // Position right below cursor
            int posX = caretPos.X;
            int posY = caretPos.Y + 20;

            // Screen boundary check
            Screen currentScreen = Screen.FromPoint(caretPos);
            if (posX + this.Width > currentScreen.WorkingArea.Right)
                posX = currentScreen.WorkingArea.Right - this.Width - 10;
            if (posY + this.Height > currentScreen.WorkingArea.Bottom)
                posY = caretPos.Y - this.Height - 5; // Position above if near bottom

            this.Location = new Point(posX, posY);
            if (!this.Visible)
            {
                this.Show();
            }
            this.Invalidate();
        }

        public string GetCandidateUnicode(int index)
        {
            if (index >= 0 && index < currentCandidates.Count)
            {
                return currentCandidates[index];
            }
            return "";
        }

        public string GetCandidateKruti(int index)
        {
            if (index >= 0 && index < currentCandidatesKruti.Count)
            {
                return currentCandidatesKruti[index];
            }
            return "";
        }

        public void Dismiss()
        {
            this.currentBuffer = "";
            this.Hide();
        }
    }

    /// <summary>
    /// Global Windows Caret Position Tracker & Keyboard Hook
    /// </summary>
    public class GlobalInputHook : IDisposable
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int X; public int Y; }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        private struct GUITHREADINFO
        {
            public int cbSize;
            public int flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public RECT rcCaret;
        }

        [DllImport("user32.dll")]
        private static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int nVirtKey);

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int KEYEVENTF_KEYUP = 0x0002;
        private const byte VK_BACK = 0x08;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private LowLevelKeyboardProc hookProc;
        private IntPtr hookId = IntPtr.Zero;

        private StringBuilder typedBuffer = new StringBuilder();
        private InlineSuggestionForm suggestionForm;
        private SpellCheckEngine spellEngine;
        private object wordApp;
        public bool IsEnabled = false;

        public GlobalInputHook(SpellCheckEngine engine, object app = null)
        {
            this.spellEngine = engine;
            this.wordApp = app;
            this.suggestionForm = new InlineSuggestionForm(engine);
            this.hookProc = HookCallback;
            this.hookId = SetHook(this.hookProc);
        }

        private IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            try
            {
                using (var curProcess = System.Diagnostics.Process.GetCurrentProcess())
                using (var curModule = curProcess.MainModule)
                {
                    IntPtr hMod = GetModuleHandle(curModule.ModuleName);
                    return SetWindowsHookEx(WH_KEYBOARD_LL, proc, hMod, 0);
                }
            }
            catch
            {
                return IntPtr.Zero;
            }
        }

        public Point GetCaretPosition()
        {
            // 1. Try Microsoft Word COM ActiveWindow.GetPoint if inside Word
            if (wordApp != null)
            {
                try
                {
                    dynamic app = wordApp;
                    if (app.Documents.Count > 0 && app.Selection != null)
                    {
                        int left = 0, top = 0, width = 0, height = 0;
                        app.ActiveWindow.GetPoint(out left, out top, out width, out height, app.Selection.Range);
                        if (left > 0 && top > 0)
                        {
                            return new Point(left, top + height);
                        }
                    }
                }
                catch { }
            }

            // 2. Fallback to Win32 GetGUIThreadInfo
            try
            {
                GUITHREADINFO guiInfo = new GUITHREADINFO();
                guiInfo.cbSize = Marshal.SizeOf(guiInfo);
                GetGUIThreadInfo(0, ref guiInfo);

                if (guiInfo.hwndCaret != IntPtr.Zero && (guiInfo.rcCaret.Left != 0 || guiInfo.rcCaret.Top != 0))
                {
                    POINT pt = new POINT { X = guiInfo.rcCaret.Left, Y = guiInfo.rcCaret.Bottom };
                    ClientToScreen(guiInfo.hwndCaret, ref pt);
                    return new Point(pt.X, pt.Y);
                }
                if (guiInfo.hwndFocus != IntPtr.Zero && (guiInfo.rcCaret.Left != 0 || guiInfo.rcCaret.Top != 0))
                {
                    POINT pt = new POINT { X = guiInfo.rcCaret.Left, Y = guiInfo.rcCaret.Bottom };
                    ClientToScreen(guiInfo.hwndFocus, ref pt);
                    return new Point(pt.X, pt.Y);
                }
            }
            catch { }

            // 3. Fallback to Cursor position
            return Cursor.Position;
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN && IsEnabled)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                Keys key = (Keys)vkCode;

                // Handle Number keys 1..5 when suggestion window is active
                if (suggestionForm.Visible && key >= Keys.D1 && key <= Keys.D5)
                {
                    int index = (int)(key - Keys.D1);
                    string selectedUni = suggestionForm.GetCandidateUnicode(index);
                    string selectedKruti = suggestionForm.GetCandidateKruti(index);
                    if (!string.IsNullOrEmpty(selectedUni))
                    {
                        ReplaceBufferWithWord(selectedUni, selectedKruti);
                        return (IntPtr)1; // Consume key
                    }
                }

                // Never auto-complete on Enter key. Enter simply dismisses the popup and passes through normally.
                if (key == Keys.Enter)
                {
                    typedBuffer.Clear();
                    suggestionForm.Dismiss();
                    return CallNextHookEx(hookId, nCode, wParam, lParam);
                }

                // Handle Tab when suggestion window is active
                if (suggestionForm.Visible && key == Keys.Tab)
                {
                    string topUni = suggestionForm.GetCandidateUnicode(0);
                    string topKruti = suggestionForm.GetCandidateKruti(0);
                    if (!string.IsNullOrEmpty(topUni))
                    {
                        ReplaceBufferWithWord(topUni, topKruti);
                        return (IntPtr)1; // Consume Tab only
                    }
                }

                // Handle Backspace
                if (key == Keys.Back)
                {
                    if (typedBuffer.Length > 0)
                    {
                        typedBuffer.Length--;
                        UpdatePopup();
                    }
                    else
                    {
                        suggestionForm.Dismiss();
                    }
                    return CallNextHookEx(hookId, nCode, wParam, lParam);
                }

                // Handle Escape (Dismiss popup)
                if (key == Keys.Escape)
                {
                    typedBuffer.Clear();
                    suggestionForm.Dismiss();
                    return CallNextHookEx(hookId, nCode, wParam, lParam);
                }

                // Handle Space or standard end-of-word punctuation
                if (key == Keys.Space || key == Keys.OemPeriod || key == Keys.Oemcomma)
                {
                    typedBuffer.Clear();
                    suggestionForm.Dismiss();
                    return CallNextHookEx(hookId, nCode, wParam, lParam);
                }

                // Capture typed characters
                char c = GetCharFromKey(key);
                if (c != '\0' && (char.IsLetterOrDigit(c) || c == '\'' || c == '\"' || c == '=' || c == '~' || c == '`' || c == ';' || c == '/' || c == '[' || c == ']' || c == '{' || c == '}'))
                {
                    typedBuffer.Append(c);
                    UpdatePopup();
                }
            }

            return CallNextHookEx(hookId, nCode, wParam, lParam);
        }

        private void UpdatePopup()
        {
            string buf = typedBuffer.ToString();
            if (buf.Length > 0)
            {
                Point caret = GetCaretPosition();
                suggestionForm.UpdateBufferAndSuggestions(buf, caret);
            }
            else
            {
                suggestionForm.Dismiss();
            }
        }

        private void ReplaceBufferWithWord(string unicodeWord, string krutiWord)
        {
            int bufLen = typedBuffer.Length;
            typedBuffer.Clear();
            suggestionForm.Dismiss();

            // Check if current Word selection is Kruti Dev or Unicode
            bool isKruti = true;
            try
            {
                if (wordApp != null)
                {
                    dynamic app = wordApp;
                    if (app.Selection != null && app.Selection.Font != null && app.Selection.Font.Name != null)
                    {
                        string fName = ((string)app.Selection.Font.Name).ToLower();
                        if (!fName.Contains("kruti") && !fName.Contains("dev 010") && !fName.Contains("walkman") && !fName.Contains("chanakya"))
                        {
                            if (fName.Contains("mangal") || fName.Contains("nirmala") || fName.Contains("aparajita") || fName.Contains("arial") || fName.Contains("calibri"))
                            {
                                isKruti = false;
                            }
                        }
                    }
                }
            }
            catch { }

            string textToSend = isKruti ? (string.IsNullOrEmpty(krutiWord) ? KrutiDevConverter.UnicodeToKruti(unicodeWord) : krutiWord) : unicodeWord;

            // Send backspaces to erase the typed prefix
            for (int i = 0; i < bufLen; i++)
            {
                keybd_event(VK_BACK, 0, 0, UIntPtr.Zero);
                keybd_event(VK_BACK, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            }

            // Send escaped text into Word
            SendKeys.SendWait(EscapeForSendKeys(textToSend));
        }

        private static string EscapeForSendKeys(string str)
        {
            if (string.IsNullOrEmpty(str)) return "";
            StringBuilder sb = new StringBuilder();
            foreach (char c in str)
            {
                if (c == '+' || c == '^' || c == '%' || c == '~' || c == '(' || c == ')' || c == '{' || c == '}')
                {
                    sb.Append("{" + c + "}");
                }
                else
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        private char GetCharFromKey(Keys key)
        {
            bool shift = (GetKeyState(0x10) & 0x8000) != 0;
            if (key >= Keys.A && key <= Keys.Z)
            {
                return shift ? key.ToString()[0] : char.ToLower(key.ToString()[0]);
            }
            if (key >= Keys.D0 && key <= Keys.D9)
            {
                return key.ToString()[1];
            }
            if (key == Keys.Oem1) return shift ? ':' : ';';
            if (key == Keys.OemQuotes) return shift ? '\"' : '\'';
            if (key == Keys.OemMinus) return shift ? '_' : '-';
            if (key == Keys.Oemplus) return shift ? '+' : '=';
            if (key == Keys.Oemtilde) return shift ? '~' : '`';
            if (key == Keys.OemOpenBrackets) return shift ? '{' : '[';
            if (key == Keys.OemCloseBrackets) return shift ? '}' : ']';
            if (key == Keys.OemQuestion) return shift ? '?' : '/';
            return '\0';
        }

        public void Dispose()
        {
            if (hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(hookId);
                hookId = IntPtr.Zero;
            }
            if (suggestionForm != null && !suggestionForm.IsDisposed)
            {
                suggestionForm.Dispose();
            }
        }
    }
}
