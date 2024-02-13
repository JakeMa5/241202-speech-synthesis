using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Speech.Synthesis;
using System.IO;
using System.Runtime.InteropServices;

namespace SpeechSynthesis
{
    public partial class TextToSpeech : Form
    {
        private SpeechSynthesizer _synthesizer = new SpeechSynthesizer();
        private string _text;
        private Form _mainForm;
        private string filePath = string.Empty;

        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        public List<String> InstalledVoices()
        {
            List<String> result = new List<String>();
            foreach (InstalledVoice voice in _synthesizer.GetInstalledVoices())
            {
                string name = voice.VoiceInfo.Name;
                result.Add(name);
            }
            return result;
        }

        public TextToSpeech(Form mainForm)
        {
            InitializeComponent();

            this.Text = string.Empty;
            this.ControlBox = false;
            this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;

            _mainForm = mainForm;

            _synthesizer.Volume = 100;
            _synthesizer.Rate = 0;
            _text = textBox1.Text;
            sliderVolume.Value = 10;

            comboBox1.DataSource = InstalledVoices();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            _synthesizer.Resume();
            _synthesizer.SpeakAsync(_text);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            _text = textBox1.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _synthesizer.SpeakAsyncCancelAll();
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            _synthesizer.Volume = sliderVolume.Value * 10;
        }

        private void sliderRate_Scroll(object sender, EventArgs e)
        {
            _synthesizer.Rate = sliderRate.Value;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            _synthesizer.SelectVoice(comboBox1.Text);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            using (var saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Wave Files (*.wav)|*.wav";
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    var outputPath = saveDialog.FileName;
                    _synthesizer.SetOutputToWaveFile(outputPath);
                    _synthesizer.Speak(_text);
                    MessageBox.Show("Audio was exported successfully.");
                    _synthesizer.SetOutputToDefaultAudioDevice();
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
            _mainForm.Show();
        }

        private void btnWindowClose_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnWindowMaximise_Click_1(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Normal)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
            }
        }

        private void btnWindowMinimise_Click_1(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
    }
}
