using System;
using System.Drawing;
using System.Windows.Forms;

namespace KrutiDevWordAddIn
{
    public class StandaloneInlineProgram
    {
        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            SpellCheckEngine engine = new SpellCheckEngine();
            GlobalInputHook hook = new GlobalInputHook(engine);

            NotifyIcon trayIcon = new NotifyIcon();
            trayIcon.Text = "कृति देव 010 इनलाइन सुझाव सहायक (Active)";
            trayIcon.Icon = SystemIcons.Application;
            trayIcon.Visible = true;

            ContextMenu menu = new ContextMenu();
            menu.MenuItems.Add("इनलाइन सुझाव चालू/बंद", (s, e) =>
            {
                hook.IsEnabled = !hook.IsEnabled;
                trayIcon.ShowBalloonTip(2000, "इनलाइन स्थिति", hook.IsEnabled ? "चालू (ON)" : "बंद (OFF)", ToolTipIcon.Info);
            });
            menu.MenuItems.Add("-");
            menu.MenuItems.Add("बंद करें (Exit)", (s, e) =>
            {
                hook.Dispose();
                trayIcon.Visible = false;
                Application.Exit();
            });

            trayIcon.ContextMenu = menu;
            trayIcon.ShowBalloonTip(3000, "कृति देव 010 इनलाइन सहायक", "इनलाइन सुझाव सक्रिय है! वर्ड या नोटपैड में टाइप करना शुरू करें।", ToolTipIcon.Info);

            Application.Run();
        }
    }
}

