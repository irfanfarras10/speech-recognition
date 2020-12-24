using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Data.SqlClient;

namespace SPEECH_ASSIST
{
    public partial class frmRegister : Form
    {
        SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine();
        SpeechSynthesizer synthesizer = new SpeechSynthesizer();


        public frmRegister()
        {
            InitializeComponent();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmLogin loginForm = new frmLogin();
            loginForm.Show();
        }

        private void loadSpeechEngine()
        {
            Choices commands = new Choices();
            commands.Add(new string[] { "login", "back to login","exit","close","register"});
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
            if (e.Result.Text == "login" || e.Result.Text == "back to login")
            {
                frmLogin loginForm = new frmLogin();
                loginForm.Show();
                this.Hide();
                recEngine.RecognizeAsyncStop();
            }
            else if (e.Result.Text == "exit" || e.Result.Text == "close")
            {
                Application.Exit();
            }
            else if (e.Result.Text == "register")
            {
                register();
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

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("You must fill data");
                synthesizer.SpeakAsync("You must fill data");
            }
            else
            {
                register();
            }
        }

        void register()
        {
            SqlConnection con = new SqlConnection("Data Source=.;Initial Catalog=SpeechAssistant;Integrated Security=True");
            con.Open();
            SqlCommand cmd = new SqlCommand("insert into [User] values(@input1,@input2,@input3)", con);
            cmd.Parameters.AddWithValue("@input1", textBox1.Text);
            cmd.Parameters.AddWithValue("@input2", textBox2.Text);
            cmd.Parameters.AddWithValue("@input3", dateTimePicker1.Value.ToString());
            cmd.ExecuteNonQuery();
            con.Close();
            synthesizer.SpeakAsync("Register Success");
            MessageBox.Show("Register Success");
            frmLogin login = new frmLogin();
            login.Show();
            this.Hide();
            recEngine.RecognizeAsyncStop();
        }
       
    }
}
