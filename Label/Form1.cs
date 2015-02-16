using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Runtime.InteropServices;

namespace Label
{
    public partial class Form1 : Form
    {
        private const int MOD_ALT = 0x1;
        private const int MOD_CONTROL = 0x2;
        private const int MOD_SHIFT = 0x4;
        private const int MOD_WIN = 0x8;
        private const int WM_HOTKEY = 0x312;
        private static AutoCompleteStringCollection state = new AutoCompleteStringCollection();
        public static OleDbConnection con = new OleDbConnection(Dataaccess.connection);
        public Form1()
        {
            InitializeComponent();
            this.BackColor = Properties.Settings.Default.Bg;
            this.ForeColor = Properties.Settings.Default.TextColour;
            this.Font = Properties.Settings.Default.Font;
            this.toolStrip1.BackColor = Properties.Settings.Default.Bg;
            this.toolStripContainer1.TopToolStripPanel.BackColor = Properties.Settings.Default.Bg;
            try
            {
                string aa = Properties.Settings.Default.imgother;
                this.BackgroundImage = ((System.Drawing.Image)(Image.FromFile(aa)));
            }
            catch { }
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (msg.Msg == 256)
            {
                if (keyData == (Keys.Escape))
                {
                    this.Close();
                }
                if (keyData == (Keys.Delete))
                {

                }
                if (ActiveControl != toolStripButton3.TextBox)
                {
                    if (keyData == (Keys.A))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.B))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.C))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.D))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.E))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.F))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.G))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.H))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.I))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.J))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.K))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.L))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.M))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.N))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.O))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.P))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.Q))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.R))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.S))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.T))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.U))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.V))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.W))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.X))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.Y))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                    else if (keyData == (Keys.Z))
                    {
                        fill_grid_KEY(keyData.ToString());
                    }
                }
            }
            //if (keyData != (Keys.Delete))
            return base.ProcessCmdKey(ref msg, keyData);
            //return base.ProcessCmdKey(ref msg, Keys.Control);
        }

        private void fill_grid_KEY(string a)
        {
            try
            {
                OleDbDataAdapter dap = new OleDbDataAdapter("select id,company as [Company Name],fname as [Name],mno as [Mobile No],phoneoffice as [Office Ph],phonehome as [Home Ph],fax,address1 as [Address],city,state,country,pin as [Zip],emailhome as [Email Home],website as [Web address],category,emailbusiness as [Email Business],profession,spouse,remarks,annvdate as [Aniversity],bdate as [B'day] from address where fname like '" + a + "%' order by fname asc", con);
                DataTable dt = new DataTable();
                if (con.State == ConnectionState.Closed) { con.Open(); }

                dap.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Visible = false;
            }
            catch { }
        }
        [DllImport("user32")]
        public static extern int RegisterHotKey(IntPtr hwnd, int id, int fsModifiers, int vk);
        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == WM_HOTKEY)
            {
                if (!Visible)
                    Visible = true;
                Activate();
                Keys vk = (Keys)(((int)m.LParam >> 16) & 0xFFFF);
                int fsModifiers = ((int)m.LParam & 0xFFFF);

                if (vk == Keys.F1 && fsModifiers == 0)
                {
                    toolStripTextBox1.TextBox.Clear();
                    fill_grid();
                }
                if (vk == Keys.F3 && fsModifiers == 0)
                {
                    this.ActiveControl = toolStripButton3.TextBox;
                }
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                RegisterHotKey(Handle, 42, 0, (int)Keys.F1);
                RegisterHotKey(Handle, 42, 0, (int)Keys.F3);
                fill_Category();
                fill_grid();
                state.Insert(0, "");
                //toolStripDropDownButton1.ComboBox.Items.Insert(0, "");
                toolStripDropDownButton1.ComboBox.SelectedIndex = -1;
                toolStripComboBox1.ComboBox.SelectedIndex = 1;
            }
            catch
            {
                MessageBox.Show("Datebase Not Found Make Sure..", "Error!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Application.Exit();
            }
        }

        private void fill_grid()
        {
            OleDbDataAdapter dap = new OleDbDataAdapter("select id,company as [Company Name],fname as [Name],mno as [Mobile No],phoneoffice as [Office Ph],phonehome as [Home Ph],fax,address1 as [Address],city,state,country,pin as [Zip],emailhome as [Email Home],website as [Web address],category,emailbusiness as [Email Business],profession,spouse,remarks,annvdate as [Aniversity],bdate as [B'day] from address", con);
            DataTable dt = new DataTable();
            if (con.State == ConnectionState.Closed) { con.Open(); }
            dap.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            Settings s = new Settings();
            s.ShowDialog();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            New_Entry n = new New_Entry();
            n.ShowDialog();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (toolStripButton3.Text == "Search")
            {
                toolStripButton3.Text = "";
            }

            else { toolStripButton3.SelectAll(); }
        }

        private void toolStripButton3_TextChanged(object sender, EventArgs e)
        {
            string aaa = "";
            string Start = "", Last = "";
            if (toolStripComboBox1.ComboBox.Text.ToString() != "-Select By-")//
            {
                if (toolStripComboBox1.ComboBox.Text == "From First") { Last = "%"; }
                if (toolStripComboBox1.ComboBox.Text == "From Last") { Start = "%"; }
                if (toolStripComboBox1.ComboBox.Text == "From AnyWhere") { Start = "%"; Last = "%"; }
            }
            if (toolStripDropDownButton1.Text != "" && toolStripDropDownButton1.Text != "Category") { aaa = " And category like '%" + toolStripDropDownButton1.Text + "%'"; }
            if (toolStripButton3.TextBox.Text != "" && toolStripButton3.TextBox.Text != "Search")
            {
                try
                {
                    OleDbDataAdapter dap = new OleDbDataAdapter("select id,company as [Company Name],fname as [Name],mno as [Mobile No],phoneoffice as [Office Ph],phonehome as [Home Ph],fax,address1 as [Address],city,state,country,pin as [Zip],emailhome as [Email Home],website as [Web address],category,emailbusiness as [Email Business],profession,spouse,remarks,annvdate as [Aniversity],bdate as [B'day] from address where fname like '" + Start + toolStripButton3.TextBox.Text.Trim() + Last + "' " + aaa + " order by fname asc", con);
                    DataTable dt = new DataTable();
                    if (con.State == ConnectionState.Closed) { con.Open(); }
                    dap.Fill(dt);
                    dataGridView1.DataSource = dt;
                    dataGridView1.Columns[0].Visible = false;
                }
                catch { }
            }
            else
            {
                if (toolStripButton3.TextBox.Text == "") { toolStripButton3.TextBox.SelectedText = "Search"; }
            }
        }

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            bool a = UnregisterHotKey(Handle, 42);
        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (e.RowIndex >= 0)
                {
                    New_Entry nw = new New_Entry(Convert.ToInt32(dataGridView1[0, e.RowIndex].Value.ToString()));
                    nw.ShowDialog();
                    toolStripButton3.TextBox.Clear();
                    fill_grid();
                }
            }
            catch { }
        }

        private void toolStripTextBox1_TextChanged(object sender, EventArgs e)
        {
            string aaa = "";
            if (toolStripDropDownButton1.Text != "" && toolStripDropDownButton1.Text != "Category") { aaa = " And category like '%" + toolStripDropDownButton1.Text + "%'"; }
            if (toolStripTextBox1.TextBox.Text != "" && toolStripTextBox1.TextBox.Text != "Mobile")
            {
                try
                {
                    OleDbDataAdapter dap = new OleDbDataAdapter("select id,company as [Company Name],fname as [Name],mno as [Mobile No],phoneoffice as [Office Ph],phonehome as [Home Ph],fax,address1 as [Address],city,state,country,pin as [Zip],emailhome as [Email Home],website as [Web address],category,emailbusiness as [Email Business],profession,spouse,remarks,annvdate as [Aniversity],bdate as [B'day] from address where (mno like '%" + toolStripTextBox1.TextBox.Text.Trim() + "%' or  mno like '%" + toolStripTextBox1.TextBox.Text.Trim() + "%' or  phoneoffice like '%" + toolStripTextBox1.TextBox.Text.Trim() + "%' or  phonehome like '%" + toolStripTextBox1.TextBox.Text.Trim() + "%' or fax like '%" + toolStripTextBox1.TextBox.Text.Trim() + "%' or centrax like '%" + toolStripTextBox1.TextBox.Text.Trim() + "%') " + aaa + "  order by fname asc", con);
                    DataTable dt = new DataTable();
                    if (con.State == ConnectionState.Closed) { con.Open(); }
                    dap.Fill(dt);
                    dataGridView1.DataSource = dt;
                    dataGridView1.Columns[0].Visible = false;
                }
                catch { }
            }
            else
            {
                if (toolStripTextBox1.TextBox.Text == "") { toolStripTextBox1.TextBox.Text = "Mobile"; }
            }
        }


        private void toolStripTextBox1_Enter(object sender, EventArgs e)
        {

        }



        private void toolStripButton3_KeyDown(object sender, KeyEventArgs e)
        {
            if (toolStripButton3.TextBox.Text == "Search")
            {
                toolStripButton3.TextBox.SelectAll();
                //toolStripButton3.TextBox.Text = e.KeyData.ToString();
            }
        }

        private void toolStripTextBox1_Click(object sender, EventArgs e)
        {
            if (toolStripTextBox1.Text == "Mobile")
            {
                toolStripTextBox1.Text = "";
            }

            else { toolStripTextBox1.SelectAll(); }
        }

        private void toolStripTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (toolStripTextBox1.TextBox.Text == "Mobile")
            {
                toolStripTextBox1.TextBox.SelectAll();
                //toolStripTextBox1.TextBox.Text = e.KeyValue.ToString();
            }
        }

        private void toolStripDropDownButton1_Click(object sender, EventArgs e)
        {

            if (toolStripDropDownButton1.ComboBox.Text == "Category")
            {
                toolStripDropDownButton1.ComboBox.Text = "";
            }
            else { toolStripDropDownButton1.ComboBox.SelectAll(); }

        }

        private void toolStripDropDownButton1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (toolStripDropDownButton1.ComboBox.Text != "" && toolStripDropDownButton1.ComboBox.Text != "Category")
                {
                    try
                    {
                        OleDbDataAdapter dap = new OleDbDataAdapter("select id,company as [Company Name],fname as [Name],mno as [Mobile No],phoneoffice as [Office Ph],phonehome as [Home Ph],fax,address1 as [Address],city,state,country,pin as [Zip],emailhome as [Email Home],website as [Web address],category,emailbusiness as [Email Business],profession,spouse,remarks,annvdate as [Aniversity],bdate as [B'day] from address where category like '%" + toolStripDropDownButton1.ComboBox.Text.Trim() + "%' order by fname asc", con);
                        DataTable dt = new DataTable();
                        if (con.State == ConnectionState.Closed) { con.Open(); }
                        dap.Fill(dt);
                        dataGridView1.DataSource = dt;
                        dataGridView1.Columns[0].Visible = false;
                    }
                    catch { }
                }
                else
                {
                    if (toolStripDropDownButton1.ComboBox.Text == "") { toolStripDropDownButton1.ComboBox.Text = "Category"; }
                }
            }
            catch { }
        }

        private void toolStripDropDownButton1_KeyDown(object sender, KeyEventArgs e)
        {
            if (toolStripDropDownButton1.Text == "Category")
            {
                toolStripDropDownButton1.ComboBox.SelectAll();
                //toolStripTextBox1.TextBox.Text = e.KeyValue.ToString();
            }
        }

        private void toolStripDropDownButton1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (toolStripDropDownButton1.ComboBox.Text != "" && toolStripDropDownButton1.ComboBox.Text != "Category")
                {
                    toolStripDropDownButton1_SelectedIndexChanged(sender, e);
                }
                else
                {
                    if (toolStripDropDownButton1.ComboBox.Text == "") { toolStripDropDownButton1.ComboBox.Text = "Category"; }
                }
            }
            catch { }
        }
        private void fill_Category()
        {
            //OleDbCommand cmd = new OleDbCommand("select distinct category from address where category<>''", con);
            //OleDbDataReader data;
            //if (con.State == ConnectionState.Closed) { con.Open(); }
            //data = cmd.ExecuteReader();
            //if (data.HasRows == true)
            //{
            //    while (data.Read())
            //    {
            //        state.Add(data["category"].ToString());
            //    }
            //}
            //data.Close();
            //toolStripDropDownButton1.ComboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            //toolStripDropDownButton1.ComboBox.AutoCompleteSource = AutoCompleteSource.CustomSource;
            //toolStripDropDownButton1.ComboBox.AutoCompleteCustomSource = state;
            //---------------------------------------------------------------------------------------------------------------------
            OleDbDataAdapter dap = new OleDbDataAdapter("select distinct category from address where category<>''", con);
            DataTable dt = new DataTable();
            if (con.State == ConnectionState.Closed) { con.Open(); }
            dap.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                toolStripDropDownButton1.ComboBox.Items.Insert(0, "Category");
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    //toolStripDropDownButton1.ComboBox.DataSource = dt;
                    //toolStripDropDownButton1.ComboBox.DisplayMember = "category";
                    toolStripDropDownButton1.ComboBox.Items.Add(dt.Rows[i][0].ToString());

                }
                toolStripDropDownButton1.ComboBox.DisplayMember = "Category";
            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            fill_grid();

        }

        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripComboBox1_TextChanged(object sender, EventArgs e)
        {
            if (toolStripComboBox1.ComboBox.Text.ToString() != "-Select By-")//
            {
                toolStripButton3.TextBox.Text = "";
            }
        }

    }
}
