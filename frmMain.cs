using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Data.SqlClient;
using System.Runtime.InteropServices;

namespace SPEECH_ASSIST
{
    public partial class frmMain : Form
    {
        SqlConnection con = new SqlConnection("Data Source=.;Initial Catalog=SpeechAssistant;Integrated Security=True");
        SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine();
        SpeechSynthesizer synthesizer = new SpeechSynthesizer();

        string username = "";

        public frmMain()
        {
            InitializeComponent();
        }

        void recommendationSystem()
        {
            con.Open();
            SqlDataReader sdr;
            SqlCommand cmd = new SqlCommand("select count(*) from Command group by CommandName", con);
            sdr = cmd.ExecuteReader();
            sdr.Read();
            int totalAccess = sdr.GetInt32(0);
            con.Close();
            if (totalAccess >= 1)
            {

                con.Open();
                SqlDataReader sdar;
                SqlCommand cmdu = new SqlCommand("select CommandName, sum(datediff(second, getdate(), AccessTime ))/count(datediff(second, getdate(), AccessTime )) from command group by CommandName", con);
                sdar = cmdu.ExecuteReader();   
                int AccessTimeTotal;
                    while (sdar.Read())
                    {
                        AccessTimeTotal = sdar.GetInt32(1);
                        if (AccessTimeTotal < 180000)
                        {
                            DialogResult result = MessageBox.Show(sdar[0].ToString(), "Recommended Command", MessageBoxButtons.OKCancel);
                            if (result == DialogResult.OK)
                            {
                                manual(sdar[0].ToString());
                            }
                        }

                    }
                con.Close();
            }

        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            recommendationSystem();
            SqlDataAdapter sda = new SqlDataAdapter("select Username from [User]",con);
            DataTable dt = new DataTable(); //this is creating a virtual table  
            sda.Fill(dt);
            username = dt.Rows[0][0].ToString();
            synthesizer.SpeakAsync("hello" + username + ".... i am your computer assistant, you can ask me and give me command");           
            richTextBox1.Text = "\nComputer : Hello " + username + "... i am your computer assistant, you can ask me and give me command";
            label4.Text = "User : " + username;
        }

        private void loadSpeechEngine()
        {
            Choices commands = new Choices();
            commands.Add(new string[] {"exit program",
                "maximize windows",
                "minimize windows",
                "show desktop",
                "who are you",
                "what time is it",
                "what is the date now",
                "what is the day now",
                "how old am i",
                "open google search",
                "open my facebook",
                "open youtube",
                "open bing search",
                "open notepad",
                "open spotify",
                "open microsoft word",
                "open microsoft excel",
                "open microsoft power point",
                "open start menu",
                "open microsoft explorer",
                "check my computer status",
                "read this text",
                "show me weather",
                "look drive D",
                "look drive C",
                "look document",
                "shutdown my pc",
                "open file manager",
                "restart my pc",
                "sleep my pc",
                "mute volume",
                "unmute volume",
                "next song",
                "previous song",
                "play song",
                "lock my pc"
            });
            GrammarBuilder gBuilder = new GrammarBuilder();
            gBuilder.Append(commands);
            Grammar grammar = new Grammar(gBuilder);

            recEngine.LoadGrammarAsync(grammar);
            recEngine.SetInputToDefaultAudioDevice();

            recEngine.SpeechRecognized += recEngine_SpeechRecognized;

            recEngine.RecognizeAsync(RecognizeMode.Multiple);
        }

        void manual(string commander)
        {
            switch (commander)
            {
                case "exit program":
                    richTextBox2.Text += "\n" + username + ": Exit";
                    synthesizer.SpeakAsync("Thank you for using me");
                    richTextBox1.Text += "\nComputer : Thank you for using me";
                    MessageBox.Show("Thank you for using me");
                    Application.Exit();
                    break;
                case "maximize windows":
                    richTextBox2.Text += "\n" + username + ": Maximize Windows";
                    richTextBox1.Text += "\nComputer : OK " + username + ", Maximizing your Windows";
                    synthesizer.SpeakAsync("OK" + username + ", Maximizing Your Windows");
                    this.WindowState = FormWindowState.Normal;
                    this.Focus();
                    break;
                case "minimize windows":
                    richTextBox2.Text += "\n" + username + ": Minimize Windows";
                    richTextBox1.Text += "\nComputer : OK " + username + ", Minimizing your Windows";
                    synthesizer.SpeakAsync("OK" + username + ", Minimizing Your Windows");
                    this.WindowState = FormWindowState.Minimized;
                    break;
                case "show desktop":
                    richTextBox2.Text += "\n" + username + ": Show Desktop";
                    richTextBox1.Text += "\nComputer : It is your desktop";
                    synthesizer.SpeakAsync("it is your desktop");
                    Shell32.ShellClass objShel = new Shell32.ShellClass();
                    objShel.ToggleDesktop();
                    break;
                case "who are you":
                    richTextBox2.Text += "\n" + username + ": Who are you ?";
                    richTextBox1.Text += "\nComputer : Hello, My Name Is LUNA Assistant, you can ask me anything";
                    synthesizer.SpeakAsync("Hello, My Name is LUNA Assistant, you can ask me anything");
                    break;
                case "what time is it":
                    string time = DateTime.Now.ToShortTimeString().ToString();
                    richTextBox2.Text += "\n" + username + ": What time is it ?";
                    richTextBox1.Text += "\nComputer : The time is : " + time;
                    synthesizer.SpeakAsync("The time is" + time);
                    break;
                case "what is the date now":
                    string date = DateTime.Now.ToString("dd MMMM yyyy");
                    richTextBox2.Text += "\n" + username + ": What is the date now ?";
                    richTextBox1.Text += "\nComputer : The date is : " + date;
                    synthesizer.SpeakAsync("The date is" + date);
                    break;
                case "what is the day now":
                    string day = DateTime.Now.ToString("dddd");
                    richTextBox2.Text += "\n" + username + ": What is the day now ?";
                    richTextBox1.Text += "\nComputer : The day is : " + day;
                    synthesizer.SpeakAsync("The day is" + day);
                    break;
                case "how old am i":
                    SqlDataAdapter sda = new SqlDataAdapter("select datediff(year, DateOfBirth, getdate()) from [User]", con);
                    DataTable dt = new DataTable(); //this is creating a virtual table  
                    sda.Fill(dt);
                    richTextBox2.Text += "\n" + username + ": How old am i ?";
                    richTextBox1.Text += "\nComputer : Your age is : " + dt.Rows[0][0];
                    synthesizer.SpeakAsync("Your age is" + dt.Rows[0][0]);
                    break;
                case "lock my pc":
                    richTextBox2.Text += "\n" + username + ": Lock my PC";
                    richTextBox1.Text += "\nComputer : Locking your PC";
                    synthesizer.SpeakAsync("Locking your PC");
                    LockWorkStation();
                    break;
                case "open google search":
                    richTextBox2.Text += "\n" + username + ": Open Google Search";
                    richTextBox1.Text += "\nComputer : Opening Google Search Engine";
                    synthesizer.SpeakAsync("Opening Google Search Engine");
                    System.Diagnostics.Process.Start("www.google.com");
                    break;
                case "open bing search":
                    richTextBox2.Text += "\n" + username + ": Open Bing Search";
                    richTextBox1.Text += "\nComputer : Opening Bing Search Engine";
                    synthesizer.SpeakAsync("Opening Bing Search Engine");
                    System.Diagnostics.Process.Start("www.bing.com");
                    break;
                case "open youtube":
                    richTextBox2.Text += "\n" + username + ": Open YouTube";
                    richTextBox1.Text += "\nComputer : Opening YouTube";
                    synthesizer.SpeakAsync("Opening YouTube");
                    System.Diagnostics.Process.Start("www.youtube.com");
                    break;
                case "open my facebook":
                    richTextBox2.Text += "\n" + username + ": Open Bing Search";
                    richTextBox1.Text += "\nComputer : Opening Your Facebook";
                    synthesizer.SpeakAsync("Opening Your Facebook");
                    System.Diagnostics.Process.Start("www.facebook.com");
                    break;
                case "mute volume":
                    richTextBox2.Text += "\n" + username + ": Mute Volume";
                    SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_MUTE);
                    break;
                case "unmute volume":
                    richTextBox2.Text += "\n" + username + ": Unmute Volume";
                    SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_UP);
                    break;
                case "next song":
                    richTextBox2.Text += "\n" + username + ": Next Song";
                    keybd_event(VK_MEDIA_NEXT_TRACK, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                    break;
                case "play song":
                    richTextBox2.Text += "\n" + username + ": Play Song";
                    keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                    break;
                case "previous song":
                    richTextBox2.Text += "\n" + username + ": Previous Song";
                    keybd_event(VK_MEDIA_PREV_TRACK, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                    break;
                case "open spotify":
                    richTextBox2.Text += "\n" + username + ": Open Spotify";
                    richTextBox1.Text += "\nComputer : Opening Your Spotify";
                    synthesizer.SpeakAsync("Opening Your Spotify");
                    System.Diagnostics.Process.Start("C:/Users/Irfan Farras/AppData/Roaming/Spotify/Spotify.exe");
                    break;
                case "open notepad":
                    richTextBox2.Text += "\n" + username + ": Open Notepad";
                    richTextBox1.Text += "\nComputer : Opening Your Notepad";
                    synthesizer.SpeakAsync("Opening Your Notepad");
                    System.Diagnostics.Process.Start("notepad");
                    break;
                case "open microsoft word":
                    richTextBox2.Text += "\n" + username + ": Open Microsoft Word";
                    richTextBox1.Text += "\nComputer : Opening Your Microsoft Word";
                    synthesizer.SpeakAsync("Opening Your Microsoft Word");
                    System.Diagnostics.Process.Start("C:/Program Files (x86)/Microsoft Office/root/Office16/WINWORD.EXE");
                    break;
                case "open microsoft excel":
                    richTextBox2.Text += "\n" + username + ": Open Microsoft Excel";
                    richTextBox1.Text += "\nComputer : Opening Your Microsoft Excel";
                    synthesizer.SpeakAsync("Opening Your Microsoft Word");
                    System.Diagnostics.Process.Start("C:/Program Files (x86)/Microsoft Office/root/Office16/EXCEL.EXE");
                    break;
                case "open microsoft power point":
                    richTextBox2.Text += "\n" + username + ": Open Microsoft Power Point";
                    richTextBox1.Text += "\nComputer : Opening Your Microsoft Power Point";
                    synthesizer.SpeakAsync("Opening Your Microsoft Power Point");
                    System.Diagnostics.Process.Start("C:/Program Files (x86)/Microsoft Office/root/Office16/POWERPNT.EXE");
                    break;
                case "open start menu":
                    richTextBox2.Text += "\n" + username + ": Open Start Menu";
                    richTextBox1.Text += "\nComputer : Opening Start Menu";
                    synthesizer.SpeakAsync("Opening Your Start Menu");
                    System.Diagnostics.Process.Start("C:/Users/Irfan Farras/Desktop/start.vbs");
                    break;
                case "open file manager":
                    richTextBox2.Text += "\n" + username + ": Open File Manager";
                    richTextBox1.Text += "\nComputer : Opening File Manager";
                    synthesizer.SpeakAsync("Opening File Manager");
                    System.Diagnostics.Process.Start("explorer");
                    break;
                case "look drive D":
                    richTextBox2.Text += "\n" + username + ": Look Drive D";
                    richTextBox1.Text += "\nComputer : It is your drive D";
                    synthesizer.SpeakAsync("It is your drive D");
                    System.Diagnostics.Process.Start("D://");
                    break;
                case "look drive C":
                    richTextBox2.Text += "\n" + username + ": Look Drive C";
                    richTextBox1.Text += "\nComputer : It is your drive C";
                    synthesizer.SpeakAsync("It is your drive C");
                    System.Diagnostics.Process.Start("C://");
                    break;
                case "look document":
                    richTextBox2.Text += "\n" + username + ": Look Drive C";
                    richTextBox1.Text += "\nComputer : It is your drive C";
                    synthesizer.SpeakAsync("It is your drive C");
                    System.Diagnostics.Process.Start("C://Users/Irfan Farras/Documents");
                    break;
                case "show me weather":
                    richTextBox2.Text += "\n" + username + ": Show  Me Weather";
                    richTextBox1.Text += "\nComputer : Looking Up For Weather";
                    synthesizer.SpeakAsync("Looking Up For Weather");
                    System.Diagnostics.Process.Start("www.google.com/search?site=&source=hp&q=weather");
                    break;
                case "read this text":
                    synthesizer.SpeakAsync(richTextBox3.Text);
                    break;
                case "check my computer status":
                    richTextBox2.Text += "\n" + username + ": Show  Me Weather";
                    richTextBox1.Text += "\nComputer : Looking Up For Weather";
                    synthesizer.SpeakAsync("Looking Up For Weather");
                    System.Diagnostics.Process.Start("%ProgramFiles%/Windows Defender/MSASCui.exe");
                    break;
            }
        }

        void recEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            switch (e.Result.Text)
            {
                case "exit program":
                    richTextBox2.Text += "\n" + username + ": Exit";
                    synthesizer.SpeakAsync("Thank you for using me");
                    richTextBox1.Text += "\nComputer : Thank you for using me";
                    MessageBox.Show("Thank you for using me");
                    Application.Exit();
                    break;
                case "maximize windows":
                    richTextBox2.Text += "\n" + username + ": Maximize Windows";
                    richTextBox1.Text += "\nComputer : OK " + username + ", Maximizing your Windows";
                    synthesizer.SpeakAsync("OK" + username + ", Maximizing Your Windows");
                    this.WindowState = FormWindowState.Normal;
                    this.Focus();
                    break;
                case "minimize windows":
                    richTextBox2.Text += "\n" + username + ": Minimize Windows";
                    richTextBox1.Text += "\nComputer : OK " + username + ", Minimizing your Windows";
                    synthesizer.SpeakAsync("OK" + username + ", Minimizing Your Windows");
                    this.WindowState = FormWindowState.Minimized;
                    break;
                case "show desktop":
                    richTextBox2.Text += "\n" + username + ": Show Desktop";
                    richTextBox1.Text += "\nComputer : It is your desktop";
                    synthesizer.SpeakAsync("it is your desktop");
                    Shell32.ShellClass objShel = new Shell32.ShellClass();
                    objShel.ToggleDesktop();
                    break;
                case "who are you" :
                    richTextBox2.Text += "\n" + username + ": Who are you ?";
                    richTextBox1.Text += "\nComputer : Hello, My Name Is LUNA Assistant, you can ask me anything";
                    synthesizer.SpeakAsync("Hello, My Name is LUNA Assistant, you can ask me anything");
                    break;
                case "what time is it":
                    string time = DateTime.Now.ToShortTimeString().ToString();
                    richTextBox2.Text += "\n" + username + ": What time is it ?";
                    richTextBox1.Text += "\nComputer : The time is : " + time;
                    synthesizer.SpeakAsync("The time is" + time);
                    break;
                case "what is the date now":
                    string date = DateTime.Now.ToString("dd MMMM yyyy");
                    richTextBox2.Text += "\n" + username + ": What is the date now ?";
                    richTextBox1.Text += "\nComputer : The date is : " + date;
                    synthesizer.SpeakAsync("The date is" + date);
                    break;
                case "what is the day now":
                    string day = DateTime.Now.ToString("dddd");
                    richTextBox2.Text += "\n" + username + ": What is the day now ?";
                    richTextBox1.Text += "\nComputer : The day is : " + day;
                    synthesizer.SpeakAsync("The day is" + day);
                    break;
                case "how old am i":
                    SqlDataAdapter sda = new SqlDataAdapter("select datediff(year, DateOfBirth, getdate()) from [User]", con);
                    DataTable dt = new DataTable(); //this is creating a virtual table  
                    sda.Fill(dt);
                    richTextBox2.Text += "\n" + username + ": How old am i ?";
                    richTextBox1.Text += "\nComputer : Your age is : " + dt.Rows[0][0];
                    synthesizer.SpeakAsync("Your age is" + dt.Rows[0][0]);
                    break;
                case "lock my pc":
                    richTextBox2.Text += "\n" + username + ": Lock my PC";
                    richTextBox1.Text += "\nComputer : Locking your PC";
                    synthesizer.SpeakAsync("Locking your PC");
                    LockWorkStation();
                    break;
                case "open google search":
                    richTextBox2.Text += "\n" + username + ": Open Google Search";
                    richTextBox1.Text += "\nComputer : Opening Google Search Engine";
                    synthesizer.SpeakAsync("Opening Google Search Engine");
                    System.Diagnostics.Process.Start("www.google.com");
                    break;
                case "open bing search":
                    richTextBox2.Text += "\n" + username + ": Open Bing Search";
                    richTextBox1.Text += "\nComputer : Opening Bing Search Engine";
                    synthesizer.SpeakAsync("Opening Bing Search Engine");
                    System.Diagnostics.Process.Start("www.bing.com");
                    break;
                case "open youtube":
                    richTextBox2.Text += "\n" + username + ": Open YouTube";
                    richTextBox1.Text += "\nComputer : Opening YouTube";
                    synthesizer.SpeakAsync("Opening YouTube");
                    System.Diagnostics.Process.Start("www.youtube.com");
                    break;
                case "open my facebook":
                    richTextBox2.Text += "\n" + username + ": Open Bing Search";
                    richTextBox1.Text += "\nComputer : Opening Your Facebook";
                    synthesizer.SpeakAsync("Opening Your Facebook");
                    System.Diagnostics.Process.Start("www.facebook.com");
                    break;
                case "mute volume":
                    richTextBox2.Text += "\n" + username + ": Mute Volume";
                    SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_MUTE);
                    break;
                case "unmute volume":
                    richTextBox2.Text += "\n" + username + ": Unmute Volume";
                    SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_UP);
                    break;
                case "next song":
                    richTextBox2.Text += "\n" + username + ": Next Song";
                    keybd_event(VK_MEDIA_NEXT_TRACK, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                    break;
                case "play song":
                    richTextBox2.Text += "\n" + username + ": Play Song";
                    keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                    break;
                case "previous song":
                    richTextBox2.Text += "\n" + username + ": Previous Song";
                    keybd_event(VK_MEDIA_PREV_TRACK, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                    break;
                case "open spotify":
                    richTextBox2.Text += "\n" + username + ": Open Spotify";
                    richTextBox1.Text += "\nComputer : Opening Your Spotify";
                    synthesizer.SpeakAsync("Opening Your Spotify");
                    System.Diagnostics.Process.Start("C:/Users/Irfan Farras/AppData/Roaming/Spotify/Spotify.exe");
                    break;
                case "open notepad":
                    richTextBox2.Text += "\n" + username + ": Open Notepad";
                    richTextBox1.Text += "\nComputer : Opening Your Notepad";
                    synthesizer.SpeakAsync("Opening Your Notepad");
                    System.Diagnostics.Process.Start("notepad");
                    break;
                case "open microsoft word":
                    richTextBox2.Text += "\n" + username + ": Open Microsoft Word";
                    richTextBox1.Text += "\nComputer : Opening Your Microsoft Word";
                    synthesizer.SpeakAsync("Opening Your Microsoft Word");
                    System.Diagnostics.Process.Start("C:/Program Files (x86)/Microsoft Office/root/Office16/WINWORD.EXE");
                    break;
                case "open microsoft excel":
                    richTextBox2.Text += "\n" + username + ": Open Microsoft Excel";
                    richTextBox1.Text += "\nComputer : Opening Your Microsoft Excel";
                    synthesizer.SpeakAsync("Opening Your Microsoft Word");
                    System.Diagnostics.Process.Start("C:/Program Files (x86)/Microsoft Office/root/Office16/EXCEL.EXE");
                    break;
                case "open microsoft power point":
                    richTextBox2.Text += "\n" + username + ": Open Microsoft Power Point";
                    richTextBox1.Text += "\nComputer : Opening Your Microsoft Power Point";
                    synthesizer.SpeakAsync("Opening Your Microsoft Power Point");
                    System.Diagnostics.Process.Start("C:/Program Files (x86)/Microsoft Office/root/Office16/POWERPNT.EXE");
                    break;
                case "open start menu":
                    richTextBox2.Text += "\n" + username + ": Open Start Menu";
                    richTextBox1.Text += "\nComputer : Opening Start Menu";
                    synthesizer.SpeakAsync("Opening Your Start Menu");
                    System.Diagnostics.Process.Start("C:/Users/Irfan Farras/Desktop/start.vbs");
                    break;
                case "open file manager":
                    richTextBox2.Text += "\n" + username + ": Open File Manager";
                    richTextBox1.Text += "\nComputer : Opening File Manager";
                    synthesizer.SpeakAsync("Opening File Manager");
                    System.Diagnostics.Process.Start("explorer");
                    break;
                case "look drive D":
                    richTextBox2.Text += "\n" + username + ": Look Drive D";
                    richTextBox1.Text += "\nComputer : It is your drive D";
                    synthesizer.SpeakAsync("It is your drive D");
                    System.Diagnostics.Process.Start("D://"); 
                    break;
                case "look drive C":
                    richTextBox2.Text += "\n" + username + ": Look Drive C";
                    richTextBox1.Text += "\nComputer : It is your drive C";
                    synthesizer.SpeakAsync("It is your drive C");
                    System.Diagnostics.Process.Start("C://");
                    break;
                case "look document":
                    richTextBox2.Text += "\n" + username + ": Look Drive C";
                    richTextBox1.Text += "\nComputer : It is your drive C";
                    synthesizer.SpeakAsync("It is your drive C");
                    System.Diagnostics.Process.Start("C://Users/Irfan Farras/Documents");
                    break;
                case "show me weather":
                    richTextBox2.Text += "\n" + username + ": Show  Me Weather";
                    richTextBox1.Text += "\nComputer : Looking Up For Weather";
                    synthesizer.SpeakAsync("Looking Up For Weather");
                    System.Diagnostics.Process.Start("www.google.com/search?site=&source=hp&q=weather");
                    break;
                case "read this text":
                    synthesizer.SpeakAsync(richTextBox3.Text);
                    break;
                case "check my computer status":
                    richTextBox2.Text += "\n" + username + ": Show  Me Weather";
                    richTextBox1.Text += "\nComputer : Looking Up For Weather";
                    synthesizer.SpeakAsync("Looking Up For Weather");
                    System.Diagnostics.Process.Start("%ProgramFiles%/Windows Defender/MSASCui.exe");
                    break;
                case "shutdown my pc":
                    richTextBox2.Text += "\n" + username + ": Shutdown My PC";
                    richTextBox1.Text += "\nComputer : Shutdown Your Computer";
                    synthesizer.SpeakAsync("Shutdown Your Computer");
                    System.Diagnostics.Process.Start("shutdown", "/s /t 0");
                    break;
                case "restart my pc":
                    richTextBox2.Text += "\n" + username + ": Restart My PC";
                    richTextBox1.Text += "\nComputer : Restart Your Computer";
                    synthesizer.SpeakAsync("Restarting Your Computer");
                    System.Diagnostics.Process.Start("shutdown", "/r /t 0");
                    break;
                case "sleep my pc":
                    richTextBox2.Text += "\n" + username + ": Sleep My PC";
                    richTextBox1.Text += "\nComputer : Sleeping Your Computer";
                    synthesizer.SpeakAsync("Sleeping Your Computer");
                    SetSuspendState(false, true, true);
                    break;
            }
        }

        [DllImport("PowrProf.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool SetSuspendState(bool hiberate, bool forceCritical, bool disableWakeEvent);

        [DllImport("user32")]
        public static extern void LockWorkStation();

        private const int APPCOMMAND_VOLUME_MUTE = 0x80000;
        private const int APPCOMMAND_VOLUME_UP = 0xA0000;
        private const int WM_APPCOMMAND = 0x319;

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, IntPtr extraInfo);

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessageW(IntPtr hWnd, int Msg,
            IntPtr wParam, IntPtr lParam);

        public const int KEYEVENTF_EXTENTEDKEY = 1;
        public const int KEYEVENTF_KEYUP = 0;
        public const int VK_MEDIA_NEXT_TRACK = 0xB0;
        public const int VK_MEDIA_PLAY_PAUSE = 0xB3;
        public const int VK_MEDIA_PREV_TRACK = 0xB1;
 

        private void button2_Click(object sender, EventArgs e)
        {
            if (button2.Text == "SPEECH : OFF")
            {
                loadSpeechEngine();
                button2.Text = "SPEECH : ON";
            }
            else
            {
                recEngine.RecognizeAsyncStop();
                button2.Text = "SPEECH : OFF";
            }
        }

    }
}
