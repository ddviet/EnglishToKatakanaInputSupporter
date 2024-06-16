using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EnglishToKatakanaInputSupporter
{
    public partial class frmMain : Form
    {
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        private const int HOTKEY_ID = 1;
        private const uint MOD_CONTROL = 0x0002;
        private const uint VK_K = 0x4B;

        private Dictionary<string, string> katakanaDictionary = new Dictionary<string, string>();
        private string dictionaryFilePath = "EnglishKatakanaDictionary.csv"; // Local file path
        private IntPtr previousWindow;

        /// <summary>
        /// 
        /// </summary>
        public frmMain()
        {
            InitializeComponent();
            RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, VK_K);
            LoadDictionaryFromFile(); // Load dictionary from file on startup
            Task.Run(() => DownloadDictionaryFile()); // Run download task in background
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
        private async Task DownloadDictionaryFile()
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
            catch (Exception ex)
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
                previousWindow = GetForegroundWindow();
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
            txtInput.Text = "";
            txtInput.Focus();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtInput_TextChanged(object sender, EventArgs e)
        {
            string input = txtInput.Text.ToLower();
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
                    this.Hide();
                    SetForegroundWindow(previousWindow);
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
