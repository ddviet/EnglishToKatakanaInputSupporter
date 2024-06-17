using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text;

namespace EnglishToKatakanaInputSupporter
{
    public partial class frmMain : Form
    {
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        private const int HOTKEY_ID = 1;
        private const uint MOD_CONTROL = 0x0002;
        private const uint VK_K = 0x4B;

        private Dictionary<string, string> katakanaDictionary = new Dictionary<string, string>();
        private string dictionaryFilePath = "EnglishKatakanaDictionary.csv"; // Local file path
        private IntPtr hWnd;

        /// <summary>
        /// 
        /// </summary>
        public frmMain()
        {
            InitializeComponent();
            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, VK_K);
            LoadDictionaryFromFile(); // Load dictionary from file on startup
            System.Threading.Tasks.Task.Run(() => DownloadDictionaryFile()); // Run download task in background
        }

        /// <summary>
        /// 
        /// </summary>
        private void LoadDictionaryFromFile()
        {
            if (File.Exists(dictionaryFilePath))
            {
                string csvData = File.ReadAllText(dictionaryFilePath);
                ParseCsv(csvData);
            }
            else
            {
                MessageBox.Show("Local dictionary file not found. Downloading from server.");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private async System.Threading.Tasks.Task DownloadDictionaryFile()
        {
            string url = "https://raw.githubusercontent.com/ddviet/EnglishToKatakanaInputSupporter/master/EnglishKatakanaDictionary.csv";
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    HttpResponseMessage response = await client.GetAsync(url);
                    if (response.IsSuccessStatusCode)
                    {
                        byte[] data = await response.Content.ReadAsByteArrayAsync();
                        File.WriteAllBytes(dictionaryFilePath, data);
                        string csvData = File.ReadAllText(dictionaryFilePath);
                        ParseCsv(csvData);
                        MessageBox.Show("Dictionary updated successfully.");
                    }
                    else
                    {
                        MessageBox.Show($"Failed to download dictionary: {response.ReasonPhrase}");
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Error downloading the Katakana dictionary: " + ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="csvData"></param>
        private void ParseCsv(string csvData)
        {
            katakanaDictionary.Clear(); // Clear existing dictionary
            using (StringReader reader = new StringReader(csvData))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    var fields = line.Replace("\"", "").Split(',');
                    if (fields.Length == 2)
                    {
                        katakanaDictionary[fields[0].Trim().ToLower()] = fields[1].Trim();
                    }
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="m"></param>
        protected override void WndProc(ref Message m)
        {
            const int WM_HOTKEY = 0x0312;
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                hWnd = GetForegroundWindow();

                string windowTitle = GetWindowTitle(hWnd);

                if (windowTitle.Contains("Word"))
                {
                    GetSelectedTextFromWord();
                }
                else if (windowTitle.Contains("Excel"))
                {
                    GetSelectedTextFromExcel();
                }
                else if (windowTitle.Contains("PowerPoint"))
                {
                    GetSelectedTextFromPowerPoint();
                }
                else
                {
                    Console.WriteLine("Active window is neither Word, Excel, Outlook, nor PowerPoint.");
                }

                ShowPopup();
            }
            base.WndProc(ref m);
        }

        /// <summary>
        /// 
        /// </summary>
        private void ShowPopup()
        {
            this.Show();
            this.Activate();
            txtInput.Text = Clipboard.GetText();
            txtInput.Focus();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtInput_TextChanged(object sender, EventArgs e)
        {
            string input = txtInput.Text.Trim().ToLower();
            if (katakanaDictionary.TryGetValue(input, out string katakana))
            {
                lblResult.Text = katakana;
            }
            else
            {
                lblResult.Text = string.Empty; // Clear label if not found
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (!string.IsNullOrEmpty(lblResult.Text))
                {
                    Clipboard.SetText(lblResult.Text);

                    this.Hide();
                    SetForegroundWindow(hWnd);
                    SendKeys.SendWait(lblResult.Text);
                }
                else
                {
                    this.Hide();
                }
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                this.Hide();
            }
            else if (e.Control && e.KeyCode == Keys.C)
            {
                if (!string.IsNullOrEmpty(lblResult.Text))
                {
                    Clipboard.SetText(lblResult.Text);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="hWnd"></param>
        /// <returns></returns>
        private string GetWindowTitle(IntPtr hWnd)
        {
            StringBuilder sb = new StringBuilder(256);
            if (GetWindowText(hWnd, sb, sb.Capacity) > 0)
            {
                return sb.ToString();
            }
            return string.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        private void GetSelectedTextFromWord()
        {
            try
            {
                Microsoft.Office.Interop.Word.Application wordApp = (Microsoft.Office.Interop.Word.Application)Marshal.GetActiveObject("Word.Application");
                Microsoft.Office.Interop.Word.Selection selection = wordApp.Selection;
                string selectedText = selection.Text;

                if (!string.IsNullOrEmpty(selectedText))
                {
                    Clipboard.SetText(selectedText);
                }

                Marshal.ReleaseComObject(wordApp);
            }
            catch (COMException)
            {
                Console.WriteLine("Word application is not running.");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void GetSelectedTextFromExcel()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
                Microsoft.Office.Interop.Excel.Range selection = excelApp.Selection;
                string selectedText = selection.Text;

                if (!string.IsNullOrEmpty(selectedText))
                {
                    Clipboard.SetText(selectedText);
                }

                Marshal.ReleaseComObject(excelApp);
            }
            catch (COMException)
            {
                Console.WriteLine("Excel application is not running.");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void GetSelectedTextFromPowerPoint()
        {
            try
            {
                Microsoft.Office.Interop.PowerPoint.Application powerPointApp = (Microsoft.Office.Interop.PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                Microsoft.Office.Interop.PowerPoint.DocumentWindow activeWindow = powerPointApp.ActiveWindow;
                if (activeWindow.Selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText)
                {
                    string selectedText = activeWindow.Selection.TextRange.Text;

                    if (!string.IsNullOrEmpty(selectedText))
                    {
                        Clipboard.SetText(selectedText);
                    }
                }
                Marshal.ReleaseComObject(activeWindow);
                Marshal.ReleaseComObject(powerPointApp);
            }
            catch (COMException)
            {
                Console.WriteLine("PowerPoint application is not running.");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="e"></param>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID);
            base.OnFormClosing(e);
        }
    }
}
