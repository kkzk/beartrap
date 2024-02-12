using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Diagnostics;
using System.Text;
using System.IO;


namespace beartrap {
    internal static class Program {
        private static string windowName;
        private static string buttonName;
        private static string EVENTLEG_SOURCE = "beartrap";
        private static AutomationElement DialogElement = null;

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
                new AutomationEventHandler(HandleOpenEvent));

            EventLog.WriteEntry(EVENTLEG_SOURCE, "イベントハンドラを設定しました", EventLogEntryType.Information, 100);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Form1 form = new Form1();
            Application.Run();
        }

        static void Capture() {
            // スクリーンショットを取得するディスプレイを選択
            int screenIndex = 0; // 0はメインディスプレイ

            // スクリーンショットを取得して保存するパス
            string pictureFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            string fileName = Path.Combine(pictureFolder, "screenshot_" + DateTime.Now.ToString("yyyy-MM-dd_HHmmss") + ".png");

            // スクリーンショットを取得して保存
            Screen screen = Screen.AllScreens[screenIndex];
            Rectangle bounds = screen.Bounds;

            // Bitmapオブジェクトを作成してスクリーンショットを取得
            using (Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height))
            using (Graphics g = Graphics.FromImage(bitmap)) {
                g.CopyFromScreen(new Point(bounds.Left, bounds.Top), Point.Empty, bounds.Size);
                bitmap.Save(fileName);
            }

            EventLog.WriteEntry(EVENTLEG_SOURCE,
                $"スクリーンショットを保存しました:{fileName}",
                EventLogEntryType.Information, 111);
        }

        // UI Automation Invokeイベントハンドラ
        private static void HandleInvokeEvent(object sender, AutomationEventArgs e) {
            EventLog.WriteEntry(EVENTLEG_SOURCE, $"[{buttonName}]が押されました", EventLogEntryType.Information, 110);

            // 画面キャプチャの取得
            Capture();

            List<string> window_names = new List<string>();
            // 操作中のファイルの取得
            TreeWalker treeWalker = TreeWalker.ControlViewWalker;
            AutomationElement mainWindow = treeWalker.GetParent(DialogElement); // main window

            // MDI のクライアント領域を取得
            AutomationElement workPane =
                mainWindow.FindFirst(TreeScope.Children,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)
                    );

            // MDI 内のウインドウを取得
            Condition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
            AutomationElementCollection workWindows = workPane.FindAll(TreeScope.Children, condition);
            foreach (AutomationElement workWindow in workWindows) {
                string name = workWindow.GetCurrentPropertyValue(AutomationElement.NameProperty) as string;
                // ウインドウのタイトルが作業中のファイル名
                window_names.Add(name);
            }

            // テーブル内テキストの取得
            List<string> table_text_lines = new List<string>();

            // グリッドコントロールを取得
            var gridControl = DialogElement.FindFirst(TreeScope.Children,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataGrid));

            var tablePattern = gridControl.GetCurrentPattern(TablePattern.Pattern) as TablePattern;
            if (tablePattern != null) {
                // グリッド内の行と列を反復処理し、各セルのテキストを取得して表示
                List<string> lineItems = new List<string>();
                for (int row = 0; row < tablePattern.Current.RowCount; row++) {
                    lineItems.Clear();
                    for (int column = 0; column < tablePattern.Current.ColumnCount; column++) {
                        var cellElement = tablePattern.GetItem(row, column);
                        var text = cellElement.Current.Name;
                        lineItems.Add(text);
                        // Console.WriteLine($"Row {row}, Column {column}: {text}");
                    }
                    table_text_lines.Add(String.Join(",", lineItems));
                }
            } else {
                Console.WriteLine("Table Control Pattern not supported.");
            }

            StringBuilder sb = new StringBuilder();
            foreach (var item in window_names) { sb.AppendLine($"作業中のファイル:{item}"); }
            foreach (var item in table_text_lines) { sb.AppendLine($"トラック:{item}"); }
            var data = sb.ToString();
            EventLog.WriteEntry(EVENTLEG_SOURCE,
                $"[{buttonName}]が押されました\n{data}",
                EventLogEntryType.Warning, 111);
        }


        private static void HandleOpenEvent(object sender, AutomationEventArgs e) {
            AutomationElement sourceElement = sender as AutomationElement;
            string window_title = sourceElement.GetCurrentPropertyValue(AutomationElement.NameProperty) as string;

            Console.WriteLine($"ウインドウが表示されました:{window_title}");

            if (window_title == windowName) {
                EventLog.WriteEntry(EVENTLEG_SOURCE,
                    $"ウインドウが表示されました:{window_title}",
                    EventLogEntryType.Information, 110);

                DialogElement = sourceElement;

                // ウインドウ内の「CD作成」ボタンを取得
                Condition condition = new PropertyCondition(AutomationElement.NameProperty, buttonName);
                AutomationElement cdBurnButton = sourceElement.FindFirst(TreeScope.Descendants, condition);

                if (cdBurnButton != null) {
                    // Invokeイベントハンドラを登録
                    Automation.AddAutomationEventHandler(
                        InvokePattern.InvokedEvent,
                        cdBurnButton,
                        TreeScope.Element,
                        new AutomationEventHandler(HandleInvokeEvent));

                    EventLog.WriteEntry(EVENTLEG_SOURCE,
                        $"ハンドラを登録しました:{buttonName}",
                        EventLogEntryType.Information, 110);
                }
                else {
                    EventLog.WriteEntry(EVENTLEG_SOURCE,
                        $"NameProperty が {buttonName} であるコントロールが見つかりませんでした",
                        EventLogEntryType.Warning, 110);
                }
            }
        }
    }
}
