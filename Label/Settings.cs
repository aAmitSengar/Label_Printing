using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Label
{
    public partial class Settings : Form
    {
        public Settings()
        {
            InitializeComponent();
            this.BackColor = Properties.Settings.Default.Bg;
            this.ForeColor = Properties.Settings.Default.TextColour;
            this.Font = Properties.Settings.Default.Font;
            try
            {
                string aa = Properties.Settings.Default.imgother;
                this.BackgroundImage = ((System.Drawing.Image)(Image.FromFile(aa)));
            }
            catch { }
        }
        [STAThread]
        private void button1_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            if (colorDialog.ShowDialog() != DialogResult.Cancel)
            {
                Properties.Settings.Default.Bg = colorDialog.Color;
                this.BackColor = Properties.Settings.Default.Bg;
                Properties.Settings.Default.Save();

            }
        }
        [STAThread]
        private void button2_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            if (colorDialog.ShowDialog() != DialogResult.Cancel)
            {
                Properties.Settings.Default.TextColour = colorDialog.Color;
                this.ForeColor = Properties.Settings.Default.TextColour;
                Properties.Settings.Default.Save();
            }
        }
        [STAThread]
        private void button4_Click(object sender, EventArgs e)
        {
            FontDialog colorDialog = new FontDialog();
            if (colorDialog.ShowDialog() != DialogResult.Cancel)
            {
                Properties.Settings.Default.Font = colorDialog.Font;
                this.Font = Properties.Settings.Default.Font;
                Properties.Settings.Default.Save();
            }
        }
        [STAThread]
        private void button3_Click(object sender, EventArgs e)
        {
            Font font = new Font("Microsoft Sans Serif", 8.25f);
            Properties.Settings.Default.Font = font;
            Properties.Settings.Default.TextColour = System.Drawing.SystemColors.ControlText;
            Properties.Settings.Default.Bg = System.Drawing.SystemColors.Control;
            this.BackColor = Properties.Settings.Default.Bg;
            this.ForeColor = Properties.Settings.Default.TextColour;
            this.Font = Properties.Settings.Default.Font;
            Properties.Settings.Default.Save();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (msg.Msg == 256)
            {
                if (keyData == (Keys.Escape))
                {
                    this.Close();
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        private void Settings_Load(object sender, EventArgs e)
        {
            //textBox1.Text=Properties.Settings.Defaul
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog f = new OpenFileDialog();
            if (f.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.img = f.FileName;
                Properties.Settings.Default.Save();
                try
                {
                    string aa = Properties.Settings.Default.img;
                    this.BackgroundImage = ((System.Drawing.Image)(Image.FromFile(aa)));
                }
                catch { }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {

            OpenFileDialog f = new OpenFileDialog();
            if (f.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.imgother = f.FileName;
                Properties.Settings.Default.Save();
                try
                {
                    string aa = Properties.Settings.Default.imgother;
                    this.BackgroundImage = ((System.Drawing.Image)(Image.FromFile(aa)));
                }
                catch { }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.img = "";
            Properties.Settings.Default.Save();
            try
            {
                string aa = Properties.Settings.Default.img;
                this.BackgroundImage = ((System.Drawing.Image)(Image.FromFile(aa)));
            }
            catch { }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.imgother = "";
            Properties.Settings.Default.Save();
            try
            {
                string aa = Properties.Settings.Default.imgother;
                this.BackgroundImage = ((System.Drawing.Image)(Image.FromFile(aa)));
            }
            catch { }
        }

        private void Settings_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(MessageBox.Show("you must restart your application to effect changes","Application Needs to restart",MessageBoxButtons.YesNo,MessageBoxIcon.Information)== DialogResult.Yes){
                Application.Restart();
            }
        }
    }
}
