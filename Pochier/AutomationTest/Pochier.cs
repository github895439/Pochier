using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Windows.Automation;

namespace AutomationTest
{
    public partial class Pochier : Form
    {
        // 定数
        private const string PROCESS_NAME_MEE = "Encoder";
        private const string PROCESS_NAME_OBS = "OBS";
        private const string BUTON_NAME_MEE = "StartStopButton";
        private const string BUTON_NAME_OBS = "5001";
        private Dictionary<string, string> MAP_APP_BUTTON = new Dictionary<string, string>();

        // 監視対象アプリ
        private Dictionary<string, Process> targetPrcss = new Dictionary<string, Process>();

        // ターゲットのボタンコントロール
        private InvokePattern targetButton;

        // 停止時刻
        private DateTime stopTime;

        public Pochier()
        {
            InitializeComponent();
        }

        #region ハンドラ

        private void Pochier_Load(object sender, EventArgs e)
        {
            Process[] aPrcss = Process.GetProcesses();

            // ターゲット探索ループ
            foreach (Process item in aPrcss)
            {
                string prcsName = item.ProcessName;

                // MEEか
                if ((!this.targetPrcss.ContainsKey(prcsName)) && (prcsName == PROCESS_NAME_MEE))
                {
                    this.targetPrcss.Add(item.ProcessName, item);
                }

                // OBSか
                if ((!this.targetPrcss.ContainsKey(prcsName)) && (prcsName == PROCESS_NAME_OBS))
                {
                    this.targetPrcss.Add(item.ProcessName, item);
                }

                // 両方見つけたか
                if (this.targetPrcss.Count == 2)
                {
                    break;
                }
            }

            // ターゲットがあったか
            if (this.targetPrcss.Count > 0)
            {
                // あった
                this.comboBox1.Items.AddRange(this.targetPrcss.Keys.ToArray());
                this.comboBox1.SelectedIndex = 0;
            }
            else
            {
                // なかった
                MessageBox.Show("アプリが起動されていません。");
                this.Close();
                return;
            }

            // 初期化
            this.Initialize();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // タイマ停止状態か
            if (this.button1.Text == "開始")
            {
                // 停止状態
                // タイマ開始
                if (this.StartTimer())
                {
                    this.panel1.Enabled = false;
                    this.button1.Text = "停止";
                    this.Text += " " + this.stopTime.ToString();
                }
            }
            else
            {
                // 開始状態
                // タイマ停止
                this.StopTimer();

                this.panel1.Enabled = true;
                this.button1.Text = "開始";
                this.Text = this.Name;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.targetButton.Invoke();
            this.timer1.Stop();
            this.panel1.Enabled = true;
            this.button1.Text = "開始";
            this.Text = this.Name;
        }

        #endregion

        #region メソッド

        private void Initialize()
        {
            this.MAP_APP_BUTTON.Add(PROCESS_NAME_MEE, BUTON_NAME_MEE);
            this.MAP_APP_BUTTON.Add(PROCESS_NAME_OBS, BUTON_NAME_OBS);

            // MEEをデフォにするループ
            for (int i = 0; i < this.comboBox1.Items.Count; i++)
            {
                // MEEか
                if ((string)this.comboBox1.Items[i] == PROCESS_NAME_MEE)
                {
                    this.comboBox1.SelectedIndex = i;
                    break;
                }
            }
        }

        private bool SetTargetButton()
        {
            Process targetPrcs = this.targetPrcss[(string)this.comboBox1.SelectedItem];
            AutomationElement target = AutomationElement.FromHandle(targetPrcs.MainWindowHandle);
            Condition conditions = new AndCondition(new PropertyCondition(AutomationElement.IsEnabledProperty, true),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
            AutomationElementCollection elements = target.FindAll(TreeScope.Subtree, conditions);
            AutomationElement button = null;

            // ボタン探索ループ
            foreach (AutomationElement item in elements)
            {
                string id = (string)item.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty);

                Debug.WriteLine(id);

                // ターゲットのボタンか
                if (id == this.MAP_APP_BUTTON[(string)this.comboBox1.SelectedItem])
                {
                    button = item;
                    break;
                }
            }

            // ボタンがなかったか
            if (button == null)
            {
                MessageBox.Show("ターゲットのボタンが見つかりません。");
                return false;
            }

            // OBSの場合、ボタンが録画停止ではないか
            if (((string)this.comboBox1.SelectedItem == PROCESS_NAME_OBS) && ((string)button.GetCurrentPropertyValue(AutomationElement.NameProperty) != "録画停止"))
            {
                MessageBox.Show("録画中ではありません。");
                return false;
            }

            this.targetButton = (InvokePattern)button.GetCurrentPattern(InvokePattern.Pattern);
            return true;
        }

        private void SetTimer()
        {
            TimeSpan tmp = this.stopTime - DateTime.Now;
            this.timer1.Interval = (int)tmp.TotalMilliseconds;
            this.timer1.Start();
        }

        private bool StartTimer()
        {
            this.stopTime = DateTime.Now;

            // 時間or時刻の「時」が省略されているか
            if (this.textBox1.Text == string.Empty)
            {
                this.textBox1.Text = "0";
            }

            // 経過指定か
            if (this.radioButton1.Checked)
            {
                try
                {
                    // 経過指定
                    TimeSpan tmpTime = new TimeSpan(Convert.ToInt32(this.textBox1.Text), Convert.ToInt32(this.textBox2.Text), 0);
                    this.stopTime = this.stopTime.Add(tmpTime);
                }
                catch (Exception)
                {
                    MessageBox.Show("指定の時間or時刻が誤っています。");
                    return false;
                }
            }
            else
            {
                try
                {
                    // 時刻指定
                    DateTime tmpTime = new DateTime(
                            this.stopTime.Year,
                            this.stopTime.Month,
                            this.stopTime.Day,
                            Convert.ToInt32(this.textBox1.Text),
                            Convert.ToInt32(this.textBox2.Text),
                            0);

                    // 日付を越えてるか
                    if (this.stopTime > tmpTime)
                    {
                        // 越えている
                        this.stopTime = tmpTime.AddDays(1);
                    }
                    else
                    {
                        // 越えていない
                        this.stopTime = tmpTime;
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("指定の時間or時刻が誤っています。");
                    return false;
                }
            }

            // ターゲットのボタン取得
            try
            {
                if (!this.SetTargetButton())
                {
                    return false;
                }
            }
            catch (Exception)
            {
                throw;
            }

            // タイマセット
            try
            {
                this.SetTimer();
            }
            catch (Exception)
            {
                throw;
            }

            return true;
        }

        private void StopTimer()
        {
            this.timer1.Stop();
            this.Text = this.Name;
        }

        #endregion
    }
}
