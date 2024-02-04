using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;


namespace beartrap {
    internal static class Program {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main() {
            windowName = "音楽CDの作成";
            buttonName = "CD作成(B)";

            // UI Automationのルートオブジェクトを取得
            AutomationElement rootElement = AutomationElement.RootElement;

            // Window が開くときのイベントを設定
            Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                rootElement,
                TreeScope.Descendants,
                HandleOpenEvent);

            Console.WriteLine("Invoke event handler registered. Press Enter to exit.");

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());
            ResidentIcon rm = new ResidentIcon();
            Application.Run();
        }

        private static string windowName;
        private static string buttonName;

        static void Capture() {
            // スクリーンショットを取得するディスプレイを選択
            int screenIndex = 0; // 0はメインディスプレイ

            // スクリーンショットを取得して保存するパス
            string fileName = "C:\\temp\\screenshot_" + DateTime.Now.ToString("yyyy-MM-dd_HHmmss") + ".png";

            // スクリーンショットを取得して保存
            CaptureScreen(screenIndex, fileName);

            Console.WriteLine("スクリーンショットが保存されました: " + fileName);
        }

        static void CaptureScreen(int screenIndex, string outputPath) {
            // スクリーンのサイズを取得
            Screen screen = Screen.AllScreens[screenIndex];
            Rectangle bounds = screen.Bounds;

            // Bitmapオブジェクトを作成してスクリーンショットを取得
            using (Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height))
            using (Graphics g = Graphics.FromImage(bitmap)) {
                g.CopyFromScreen(new Point(bounds.Left, bounds.Top), Point.Empty, bounds.Size);
                bitmap.Save(outputPath);
            }
        }

        // UI Automation Invokeイベントハンドラ
        private static void HandleInvokeEvent(object sender, AutomationEventArgs e) {
            Console.WriteLine("Invoke event handled.");

            // 画面キャプチャの取得
            Capture();
        }

        private static void HandleOpenEvent(object sender, AutomationEventArgs e) {
            AutomationElement sourceElement = sender as AutomationElement;
            string window_title = sourceElement.GetCurrentPropertyValue(AutomationElement.NameProperty) as string;
            Console.WriteLine("Open event handled.{0}", window_title);
            if (window_title == windowName) {
                // ウインドウ内の「CD作成」ボタンを取得
                Condition condition = new PropertyCondition(AutomationElement.NameProperty, buttonName);
                AutomationElement cdBurnButton = sourceElement.FindFirst(TreeScope.Descendants, condition);

                // Invokeイベントハンドラを登録
                Automation.AddAutomationEventHandler(
                    InvokePattern.InvokedEvent,
                    cdBurnButton,
                    TreeScope.Element,
                    HandleInvokeEvent);
            }
        }
    }

    class ResidentIcon : Form {
        public ResidentIcon() {
            this.ShowInTaskbar = false;
            this.setComponents();
        }

        private void Close_Click(object sender, EventArgs e) {
            Application.Exit();
        }

        private void setComponents() {
            NotifyIcon icon = new NotifyIcon();
            icon.Icon = new Icon("beartrap.ico");
            icon.Visible = true;
            icon.Text = "beartrap";
            ContextMenuStrip menu = new ContextMenuStrip();
            ToolStripMenuItem menuItem = new ToolStripMenuItem();
            menuItem.Text = "&終了";
            menuItem.Click += new EventHandler(Close_Click);
            menu.Items.Add(menuItem);
            icon.ContextMenuStrip = menu;
        }
    }

}
