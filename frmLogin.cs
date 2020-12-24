using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Speech.Synthesis;
using System.Speech.Recognition;


namespace SPEECH_ASSIST
{
    public partial class frmLogin : Form
    {
        SqlConnection con = new SqlConnection("Data Source=.;Initial Catalog=SpeechAssistant;Integrated Security=True");
        SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine();
        SpeechSynthesizer synthesizer = new SpeechSynthesizer();

        public frmLogin()
        {
            InitializeComponent();
        }

        private void loadSpeechEngine()
        {
            Choices commands = new Choices();
            commands.Add(new string[] { "exit program","exit","close","register","login"});
            GrammarBuilder gBuilder = new GrammarBuilder();
            gBuilder.Append(commands);
            Grammar grammar = new Grammar(gBuilder);

            recEngine.LoadGrammarAsync(grammar);
            recEngine.SetInputToDefaultAudioDevice();

            recEngine.SpeechRecognized += recEngine_SpeechRecognized;

            recEngine.RecognizeAsync(RecognizeMode.Multiple);
        }

        void recEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result.Text == "exit" || e.Result.Text == "exit program" || e.Result.Text == "close")
            {
                Application.Exit();
            }
            else if (e.Result.Text == "register")
            {
                this.Hide();
                frmRegister registerForm = new frmRegister();
                registerForm.Show();
                recEngine.RecognizeAsyncStop();
            }
            else if(e.Result.Text == "login")
            {
                login();
            }
        }

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

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmRegister registerForm = new frmRegister();
            registerForm.Show();
        }

        private void frmLogin_Load(object sender, EventArgs e)
        {
            recEngine.RecognizeAsyncStop();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            login();
        }

        void login(){
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("Username and Password must be fill");
            }
            else
            {
                SqlDataAdapter sda = new SqlDataAdapter("SELECT COUNT(*) FROM [User] WHERE Username='" + textBox1.Text + "' AND Password='" + textBox2.Text + "'", con);

                DataTable dt = new DataTable(); //this is creating a virtual table  
                sda.Fill(dt);
                if (dt.Rows[0][0].ToString() == "1")
                {
                    /* I have made a new page called home page. If the user is successfully authenticated then the form will be moved to the next form */
                    this.Hide();
                    recEngine.RecognizeAsyncStop();
                    frmMain mainForm = new frmMain();
                    mainForm.Show();
                }
                else
                {
                    synthesizer.SpeakAsync("Username or Password are Incorrect");
                    MessageBox.Show("Username or Password are Incorrect");
                } 
            }
        }
    }
}
