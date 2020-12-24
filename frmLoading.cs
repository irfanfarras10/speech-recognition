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
    public partial class frmLoading : Form
    {
        SqlConnection con = new SqlConnection("Data Source=.;Initial Catalog=SpeechAssistant;Integrated Security=True");
        SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine();
        SpeechSynthesizer synthesizer = new SpeechSynthesizer();

        public frmLoading()
        {
            InitializeComponent();
        }

        private void Loading_Load(object sender, EventArgs e)
        {
            timer1.Interval = 6000;
            timer1.Start();
            synthesizer.SpeakAsync("Please Wait ... Loading Components");
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();
            if (timer1.Interval == 6000)
            {
                if (login() == 0)
                {
                    frmLogin loginForm = new frmLogin();
                    loginForm.Show();
                    this.Hide();
                }
                else
                {
                    frmMain mainForm = new frmMain();
                    mainForm.Show();
                }
            }
        }

        int login()
        {
            SqlDataAdapter sda = new SqlDataAdapter("select count(*) from [User]", con);
            DataTable dt = new DataTable(); //this is creating a virtual table  
            sda.Fill(dt);
            if (dt.Rows[0][0].ToString() == "1")
            {
                /* I have made a new page called home page. If the user is successfully authenticated then the form will be moved to the next form */
                this.Hide();
                return 1;
            }
            else
                return 0;
        }
    }
}
