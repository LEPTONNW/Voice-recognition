using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Speech.Recognition;
using Microsoft.Speech.Synthesis;
using System.Threading;
using System.Runtime.InteropServices;
using System.IO;

namespace SR_Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            initRS();
            initTTS();
        }
        [DllImport("user32.dll")]
        public static extern void keybd_event(uint vk, uint scan, uint flags, uint extraInfo);

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        public static string FilePath = @"BigData\";
        

        public static SpeechRecognitionEngine sre = new SpeechRecognitionEngine(new CultureInfo("ko-KR"));
        public static Grammar g = new Grammar("T.xml");

        protected override void WndProc(ref Message m) //FormboardStyle = None 일 경우 마우스 제어 함수
        {
            switch (m.Msg)
            {
                case 0x84:
                    base.WndProc(ref m);
                    if ((int)m.Result == 0x1)
                        m.Result = (IntPtr)0x2;
                    return;
            }

            base.WndProc(ref m);
        }

        public void initRS()
        {
            try
            {
                sre = new SpeechRecognitionEngine(new CultureInfo("ko-KR"));
                sre.LoadGrammar(g);
                sre.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sre_SpeechRecognized);
                sre.SetInputToDefaultAudioDevice();
                sre.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch (Exception e)
            {
                label1.Text = "init RS Error : " + e.ToString();
            }
        }

        SpeechSynthesizer tts;

        public void initTTS()
        {
            try 
            {
                tts = new SpeechSynthesizer();
                //tts.SelectVoice("Microsoft Server Speech Text to Speech Voice (ko-KR, Heami)");
                tts.SelectVoice("Microsoft Server Speech Text to Speech Voice (en-US, ZiraPro)");
                tts.SetOutputToDefaultAudioDevice();
                tts.Volume = 100;
            }
            catch(Exception e)
            {
                label1.Text = "init TTS Error : " + e.ToString();
            }
        }

        void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            label1.Text = e.Result.Text;

            if (e.Result.Text == "컴퓨터")
            {
                tts.SpeakAsync("Sir?");
            }
            else if (e.Result.Text == "셧다운" || e.Result.Text == "종료")
            {
                tts.Speak("Program Exit. Good Bye");
                Application.Exit();
            }
            else if (e.Result.Text == "구글")
            {
                tts.SpeakAsync("Launch a Google website");
                System.Diagnostics.Process.Start("chrome.exe", "https://www.google.co.kr/");

                if (File.Exists(FilePath + "google.ini"))
                {
                    StringBuilder temp = new StringBuilder(255);
                    int ret = GetPrivateProfileString("구글", "사이트빈도", "", temp, 255, FilePath + "google.ini");
                    int sum = Convert.ToInt32(temp.ToString());

                    WritePrivateProfileString("구글", "사이트빈도", (sum + 1).ToString(), FilePath + "google.ini");
                }
                else
                {
                    WritePrivateProfileString("구글", "사이트빈도", "1", FilePath + "google.ini");
                }
            }
            else if (e.Result.Text == "개드립")
            {
                tts.SpeakAsync("Complete WebSite execution");
                System.Diagnostics.Process.Start("chrome.exe", "http://www.dogdrip.net/");

                if (File.Exists(FilePath + "dogdrip.ini"))
                {
                    StringBuilder temp = new StringBuilder(255);
                    int ret = GetPrivateProfileString("개드립", "사이트빈도", "", temp, 255, FilePath + "dogdrip.ini");
                    int sum = Convert.ToInt32(temp.ToString());

                    WritePrivateProfileString("개드립", "사이트빈도", (sum + 1).ToString(), FilePath + "dogdrip.ini");
                }
                else
                {
                    WritePrivateProfileString("개드립", "사이트빈도", "1", FilePath + "dogdrip.ini");
                }
            }
            else if (e.Result.Text == "네이버")
            {
                tts.SpeakAsync("Complete WebSite execution");
                System.Diagnostics.Process.Start("chrome.exe", "http://www.naver.com/");

                if (File.Exists(FilePath + "naver.ini"))
                {
                    StringBuilder temp = new StringBuilder(255);
                    int ret = GetPrivateProfileString("네이버", "사이트빈도", "", temp, 255, FilePath + "naver.ini");
                    int sum = Convert.ToInt32(temp.ToString());

                    WritePrivateProfileString("네이버", "사이트빈도", (sum + 1).ToString(), FilePath + "naver.ini");
                }
                else
                {
                    WritePrivateProfileString("네이버", "사이트빈도", "1", FilePath + "naver.ini");
                }
            }
            else if (e.Result.Text == "유튜브")
            {
                tts.SpeakAsync("Complete WebSite execution");
                System.Diagnostics.Process.Start("chrome.exe", "http://www.youtube.co.kr");

                if (File.Exists(FilePath + "youtube.ini"))
                {
                    StringBuilder temp = new StringBuilder(255);
                    int ret = GetPrivateProfileString("유튜브", "사이트빈도", "", temp, 255, FilePath + "youtube.ini");
                    int sum = Convert.ToInt32(temp.ToString());

                    WritePrivateProfileString("유튜브", "사이트빈도", (sum + 1).ToString(), FilePath + "youtube.ini");
                }
                else
                {
                    WritePrivateProfileString("유튜브", "사이트빈도", "1", FilePath + "youtube.ini");
                }
            }
            else if (e.Result.Text == "엘엠에스")
            {
                tts.SpeakAsync("Complete WebSite execution");
                System.Diagnostics.Process.Start("chrome.exe", "http://kpu.kdual.net/");

                if (File.Exists(FilePath + "LMS.ini"))
                {
                    StringBuilder temp = new StringBuilder(255);
                    int ret = GetPrivateProfileString("엘엠에스", "사이트빈도", "", temp, 255, FilePath + "LMS.ini");
                    int sum = Convert.ToInt32(temp.ToString());

                    WritePrivateProfileString("엘엠에스", "사이트빈도", (sum + 1).ToString(), FilePath + "LMS.ini");
                }
                else
                {
                    WritePrivateProfileString("엘엠에스", "사이트빈도", "1", FilePath + "LMS.ini");
                }
            }
            else if (e.Result.Text == "케이피유")
            {
                tts.SpeakAsync("Complete WebSite execution");
                System.Diagnostics.Process.Start("chrome.exe", "https://www.kpu.ac.kr/index.do");

                if (File.Exists(FilePath + "KPU.ini"))
                {
                    StringBuilder temp = new StringBuilder(255);
                    int ret = GetPrivateProfileString("케이피유", "사이트빈도", "", temp, 255, FilePath + "KPU.ini");
                    int sum = Convert.ToInt32(temp.ToString());

                    WritePrivateProfileString("케이피유", "사이트빈도", (sum + 1).ToString(), FilePath + "KPU.ini");
                }
                else
                {
                    WritePrivateProfileString("케이피유", "사이트빈도", "1", FilePath + "KPU.ini");
                }
            }
            else if (e.Result.Text == "한글")
            {
                tts.SpeakAsync("Complete program execution");
                doProgram("C:\\Program Files (x86)\\Hnc\\Hwp80\\Hwp.exe", "");
            }
            else if (e.Result.Text == "한글종료")
            {
                tts.SpeakAsync("Program Exit");
                closeProcess("Hwp");
            }
            else if (e.Result.Text == "엑셀")
            {
                tts.SpeakAsync("Complete program execution");
                doProgram("C:\\Program Files\\Microsoft Office\\Office14\\EXCEL.EXE", "");
            }
            else if (e.Result.Text == "피피티" || e.Result.Text == "파워포인트")
            {
                tts.SpeakAsync("Complete program execution");
                doProgram("C:\\Program Files\\Microsoft Office\\Office14\\POWERPNT.EXE", "");
            }
            else if (e.Result.Text == "아래로" || e.Result.Text == "아래")
            {
                keybd_event((byte)Keys.PageDown, 0, 0x00, 0);
                keybd_event((byte)Keys.PageDown, 0, 0x02, 0);
            }
            else if (e.Result.Text == "위로" || e.Result.Text == "위")
            {
                keybd_event((byte)Keys.PageUp, 0, 0x00, 0);
                keybd_event((byte)Keys.PageUp, 0, 0x02, 0);
            }
            else if (e.Result.Text == "사이트추천해줘" || e.Result.Text == "사이트" || e.Result.Text == "추천해줘")
            {
                List<int> myList = new List<int>();
                List<string> site = new List<string>();
                int max = 0;
                string count = "";

                try
                {
                    StringBuilder temp1 = new StringBuilder(255);
                    int ret1 = GetPrivateProfileString("구글", "사이트빈도", "", temp1, 255, FilePath + "google.ini");
                    myList.Add(Convert.ToInt32(temp1.ToString()));
                    site.Add("google");
                }
                catch { }

                try
                {
                    StringBuilder temp2 = new StringBuilder(255);
                    int ret2 = GetPrivateProfileString("케이피유", "사이트빈도", "", temp2, 255, FilePath + "KPU.ini");
                    myList.Add(Convert.ToInt32(temp2.ToString()));
                    site.Add("KPU");
                }
                catch { }

                try
                {
                    StringBuilder temp3 = new StringBuilder(255);
                    int ret3 = GetPrivateProfileString("엘엠에스", "사이트빈도", "", temp3, 255, FilePath + "LMS.ini");
                    myList.Add(Convert.ToInt32(temp3.ToString()));
                    site.Add("LMS");
                }
                catch { }

                try
                {
                    StringBuilder temp4 = new StringBuilder(255);
                    int ret4 = GetPrivateProfileString("개드립", "사이트빈도", "", temp4, 255, FilePath + "dogdrip.ini");
                    myList.Add(Convert.ToInt32(temp4.ToString()));
                    site.Add("dogdrip");
                }
                catch { }

                try
                {
                    StringBuilder temp5 = new StringBuilder(255);
                    int ret5 = GetPrivateProfileString("네이버", "사이트빈도", "", temp5, 255, FilePath + "naver.ini");
                    myList.Add(Convert.ToInt32(temp5.ToString()));
                    site.Add("naver");
                }
                catch { }

                try
                {
                    StringBuilder temp6 = new StringBuilder(255);
                    int ret6 = GetPrivateProfileString("유튜브", "사이트빈도", "", temp6, 255, FilePath + "youtube.ini");
                    myList.Add(Convert.ToInt32(temp6.ToString()));
                    site.Add("youtube");
                }
                catch { }

                int j = 0;
                foreach (int i in myList)
                {
                    if (i > max)
                    {
                        max = i;
                        //MessageBox.Show(site[j].ToString());
                        count = site[j];

                    }
                    j++;
                }
                try
                {
                    if (count == "google")
                    {
                        if (MessageBox.Show("가장 많이 이용한 사이트는구글 입니다." + Environment.NewLine + "접속하시겠습니까?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start("chrome.exe", "https://www.google.co.kr/");
                        }
                    }
                    else if (count == "KPU")
                    {
                        if (MessageBox.Show("가장 많이 이용한 케이피유 입니다." + Environment.NewLine + "접속하시겠습니까?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start("chrome.exe", "https://www.kpu.ac.kr/index.do");
                        }
                    }
                    else if (count == "LMS")
                    {
                        if (MessageBox.Show("가장 많이 이용한 엘엠에스 입니다." + Environment.NewLine + "접속하시겠습니까?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start("chrome.exe", "http://kpu.kdual.net/");
                        }
                    }
                    else if (count == "dogdrip")
                    {
                        if (MessageBox.Show("가장 많이 이용한 개드립 입니다." + Environment.NewLine + "접속하시겠습니까?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start("chrome.exe", "http://www.dogdrip.net");
                        }
                    }
                    else if (count == "naver")
                    {
                        if (MessageBox.Show("가장 많이 이용한 네이버 입니다." + Environment.NewLine + "접속하시겠습니까?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start("chrome.exe", "http://www.naver.com");
                        }
                    }
                    else if (count == "youtube")
                    {
                        if (MessageBox.Show("가장 많이 이용한 유튜브 입니다." + Environment.NewLine + "접속하시겠습니까?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start("chrome.exe", "http://www.youtube.co.kr");
                        }
                    }
                }
                catch
                {
                    tts.SpeakAsync("Yes, Sir?");
                    if (count == "google")
                    {
                        if (MessageBox.Show("가장 많이 이용한 사이트는구글 입니다." + Environment.NewLine + "접속하시겠습니까?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start("iexplore.exe", "https://www.google.co.kr/");
                        }
                    }
                    else if (count == "KPU")
                    {
                        if (MessageBox.Show("가장 많이 이용한 케이피유 입니다." + Environment.NewLine + "접속하시겠습니까?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start("iexplore.exe", "https://www.kpu.ac.kr/index.do");
                        }
                    }
                    else if (count == "LMS")
                    {
                        if (MessageBox.Show("가장 많이 이용한 엘엠에스 입니다." + Environment.NewLine + "접속하시겠습니까?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start("iexplore.exe", "http://kpu.kdual.net/");
                        }
                    }
                    else if (count == "dogdrip")
                    {
                        if (MessageBox.Show("가장 많이 이용한 개드립 입니다." + Environment.NewLine + "접속하시겠습니까?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start("iexplore.exe", "http://www.dogdrip.net");
                        }
                    }
                    else if (count == "naver")
                    {
                        if (MessageBox.Show("가장 많이 이용한 네이버 입니다." + Environment.NewLine + "접속하시겠습니까?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start("iexplore.exe", "http://www.naver.com");
                        }
                    }
                    else if (count == "youtube")
                    {
                        if (MessageBox.Show("가장 많이 이용한 유튜브 입니다." + Environment.NewLine + "접속하시겠습니까?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start("iexplore.exe", "http://www.youtube.co.kr");
                        }
                    }
                }
            }
            else if (e.Result.Text == "사이트빈도확인해줘" || e.Result.Text == "사이트빈도" || e.Result.Text == "확인해줘")
            {
                List<string> myList = new List<string>();
                string STsum = "";

                try
                {
                    StringBuilder temp1 = new StringBuilder(255);
                    int ret1 = GetPrivateProfileString("구글", "사이트빈도", "", temp1, 255, FilePath + "google.ini");
                    myList.Add("구글 : " + temp1.ToString());
                }
                catch { }

                try
                {
                    StringBuilder temp2 = new StringBuilder(255);
                    int ret2 = GetPrivateProfileString("케이피유", "사이트빈도", "", temp2, 255, FilePath + "KPU.ini");
                    myList.Add("KPU : " + temp2.ToString());
                }
                catch { }

                try
                {
                    StringBuilder temp3 = new StringBuilder(255);
                    int ret3 = GetPrivateProfileString("엘엠에스", "사이트빈도", "", temp3, 255, FilePath + "LMS.ini");
                    myList.Add("LMS : " + temp3.ToString());
                }
                catch { }

                try
                {
                    StringBuilder temp4 = new StringBuilder(255);
                    int ret4 = GetPrivateProfileString("개드립", "사이트빈도", "", temp4, 255, FilePath + "dogdrip.ini");
                    myList.Add("개드립 : " + temp4.ToString());
                }
                catch { }

                try
                {
                    StringBuilder temp5 = new StringBuilder(255);
                    int ret5 = GetPrivateProfileString("네이버", "사이트빈도", "", temp5, 255, FilePath + "naver.ini");
                    myList.Add("네이버 : " + temp5.ToString());
                }
                catch { }

                try
                {
                    StringBuilder temp6 = new StringBuilder(255);
                    int ret6 = GetPrivateProfileString("유튜브", "사이트빈도", "", temp6, 255, FilePath + "youtube.ini");
                    myList.Add("유튜브 : " + temp6.ToString());
                }
                catch { }

                for (int i = 0; i < 6; i++)
                {
                    STsum += myList[i].ToString() + Environment.NewLine;
                }
                tts.SpeakAsync("Yes, Sir?");
                MessageBox.Show("[사이트 빈도]" + Environment.NewLine + STsum);

            }
            else
            {
                label1.Text = e.Result.Text;
            }
            button1.Enabled = true;

        }

        // 프로세스 실행
        private static void doProgram(string filename, string arg)
        {
            ProcessStartInfo psi;
            if(arg.Length != 0)
                psi = new ProcessStartInfo(filename, arg);
            else
                psi = new ProcessStartInfo(filename);
            Process.Start(psi);
        }

        // 프로세스 종료
        private static void closeProcess(string filename)
        {
            Process[] myProcesses;
            // Returns array containing all instances of Notepad.
            myProcesses = Process.GetProcessesByName(filename);
            foreach (Process myProcess in myProcesses)
            {
                myProcess.CloseMainWindow();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if(!Directory.Exists(FilePath))
            {
                Directory.CreateDirectory(FilePath);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.button1.BackColor = System.Drawing.Color.LightGreen;
            this.button2.BackColor = System.Drawing.Color.White;
            button1.Enabled = false;
            try
            {
                sre.Dispose();
                sre = new SpeechRecognitionEngine(new CultureInfo("ko-KR"));
                sre.LoadGrammar(g);
                sre.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sre_SpeechRecognized);
                sre.SetInputToDefaultAudioDevice();
                sre.RecognizeAsync(RecognizeMode.Single);
            }
            catch (Exception w)
            {
                label1.Text = "init RS Error : " + w.ToString();
                sre.Dispose();
                sre = new SpeechRecognitionEngine(new CultureInfo("ko-KR"));
                sre.LoadGrammar(g);
                sre.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sre_SpeechRecognized);
                sre.SetInputToDefaultAudioDevice();
                sre.RecognizeAsync(RecognizeMode.Single);
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            sre.Dispose();
            initRS();
            this.button2.BackColor = System.Drawing.Color.LightGreen;
            this.button1.BackColor = System.Drawing.Color.White;
        }
    }
}
