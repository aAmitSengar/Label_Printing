using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Data.OleDb;

namespace Label
{
    public partial class New_Entry : Form
    {
        public static OleDbConnection con = new OleDbConnection(Dataaccess.connection);
        private AutoCompleteStringCollection catagry = new AutoCompleteStringCollection();
        private AutoCompleteStringCollection country = new AutoCompleteStringCollection();
        private AutoCompleteStringCollection code = new AutoCompleteStringCollection();
        private AutoCompleteStringCollection city = new AutoCompleteStringCollection();
        private AutoCompleteStringCollection state = new AutoCompleteStringCollection();
        private AutoCompleteStringCollection cmpny = new AutoCompleteStringCollection();
        private static Int32 ID = 0;
        public New_Entry(Int32 IID)
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
            ID = IID;
        }
        public New_Entry()
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
            ID = 0;
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
        private void New_Entry_Load(object sender, EventArgs e)
        {
            cmb_title.SelectedIndex = 0;
            fill_Catagory();
            fill_country();
            fill_CODE();
            fill_city();
            fill_State();
            fill_cmpny();
            cmd_ctgry.Text = "";
            txt_companyName.Text = "";
            txt_city.Text = "";
            txt_State.Text = "";
            txt_STD.Text = "";
            if (ID != 0)
            {
                button1.Text = "&Update";
                string SQL = "";
                SQL = "select title,centrax,company,fname,mno,phoneoffice,phonehome,fax,address1,city,state,country,pin,emailhome,website,category,emailbusiness,profession,spouse,remarks,annvdate,bdate,std from address where id=" + ID;
                DataTable dt = new DataTable();
                OleDbDataAdapter dap = new OleDbDataAdapter(SQL, con);
                dap.Fill(dt);
                if (dt.Rows.Count == 1)
                {
                    cmb_title.Text = dt.Rows[0]["title"].ToString();
                    cmd_ctgry.Text = dt.Rows[0]["category"].ToString();
                    txt_name.Text = dt.Rows[0]["fname"].ToString();
                    txt_companyName.Text = dt.Rows[0]["company"].ToString();
                    txt_profation.Text = dt.Rows[0]["profession"].ToString();
                    txt_supose.Text = dt.Rows[0]["spouse"].ToString();
                    txt_address.Text = dt.Rows[0]["address1"].ToString();
                    txt_city.Text = dt.Rows[0]["city"].ToString();
                    txt_State.Text = dt.Rows[0]["state"].ToString();
                    txt_STD.Text = dt.Rows[0]["std"].ToString();
                    txt_webpage.Text = dt.Rows[0]["website"].ToString();
                    txt_Zip.Text = dt.Rows[0]["pin"].ToString();
                    txt_contact.Text = dt.Rows[0]["centrax"].ToString();
                    txt_country.Text = dt.Rows[0]["country"].ToString();
                    txt_email_B.Text = dt.Rows[0]["emailbusiness"].ToString();
                    txt_email_h.Text = dt.Rows[0]["emailhome"].ToString();
                    txt_fax.Text = dt.Rows[0]["fax"].ToString();
                    txt_Home_phone.Text = dt.Rows[0]["phonehome"].ToString();
                    txt_mob.Text = dt.Rows[0]["mno"].ToString();
                    txt_phone_offc.Text = dt.Rows[0]["phoneoffice"].ToString();
                    txt_remarks.Text = dt.Rows[0]["remarks"].ToString();
                }
            }
            else { button1.Text = "&Save"; }
            ActiveControl = txt_name;
        }

        private void fill_cmpny()
        {
            OleDbCommand cmd = new OleDbCommand("select distinct company from address where company<>''", con);
            OleDbDataReader data;
            if (con.State == ConnectionState.Closed) { con.Open(); }
            data = cmd.ExecuteReader();
            if (data.HasRows == true)
            {
                while (data.Read())
                {
                    cmpny.Add(data["company"].ToString());
                }
            }
            txt_companyName.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txt_companyName.AutoCompleteSource = AutoCompleteSource.CustomSource;
            txt_companyName.AutoCompleteCustomSource = catagry;
            //---------------------------------------------------------------------------------------------------------------------
            OleDbDataAdapter dap = new OleDbDataAdapter("select distinct company from address where company<>''", con);
            DataTable dt = new DataTable();
            dap.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                txt_companyName.DataSource = dt;
                txt_companyName.DisplayMember = "company";
            }
        }

        private void fill_State()
        {
            OleDbCommand cmd = new OleDbCommand("select distinct state from address where state<>''", con);
            OleDbDataReader data;
            if (con.State == ConnectionState.Closed) { con.Open(); }
            data = cmd.ExecuteReader();
            if (data.HasRows == true)
            {
                while (data.Read())
                {
                    state.Add(data["state"].ToString());
                }
            }
            txt_State.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txt_State.AutoCompleteSource = AutoCompleteSource.CustomSource;
            txt_State.AutoCompleteCustomSource = catagry;
            //---------------------------------------------------------------------------------------------------------------------
            OleDbDataAdapter dap = new OleDbDataAdapter("select distinct state from address where state<>''", con);
            DataTable dt = new DataTable();
            dap.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                txt_State.DataSource = dt;
                txt_State.DisplayMember = "state";
            }
        }

        private void fill_city()
        {
            OleDbCommand cmd = new OleDbCommand("select distinct city from address where city<>''", con);
            OleDbDataReader data;
            if (con.State == ConnectionState.Closed) { con.Open(); }
            data = cmd.ExecuteReader();
            if (data.HasRows == true)
            {
                while (data.Read())
                {
                    city.Add(data["city"].ToString());
                }
            }
            txt_city.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txt_city.AutoCompleteSource = AutoCompleteSource.CustomSource;
            txt_city.AutoCompleteCustomSource = catagry;
            //---------------------------------------------------------------------------------------------------------------------
            OleDbDataAdapter dap = new OleDbDataAdapter("select distinct city from address where city<>''", con);
            DataTable dt = new DataTable();
            dap.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                txt_city.DataSource = dt;
                txt_city.DisplayMember = "city";
            }
        }

        private void fill_CODE()
        {
            OleDbCommand cmd = new OleDbCommand("select distinct std from address where std<>''", con);
            OleDbDataReader data;
            if (con.State == ConnectionState.Closed) { con.Open(); }
            data = cmd.ExecuteReader();
            if (data.HasRows == true)
            {
                while (data.Read())
                {
                    code.Add(data["std"].ToString());
                }
            }
            txt_STD.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txt_STD.AutoCompleteSource = AutoCompleteSource.CustomSource;
            txt_STD.AutoCompleteCustomSource = catagry;
            //---------------------------------------------------------------------------------------------------------------------
            OleDbDataAdapter dap = new OleDbDataAdapter("select distinct std from address where std<>''", con);
            DataTable dt = new DataTable();
            dap.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                txt_STD.DataSource = dt;
                txt_STD.DisplayMember = "std";
            }
        }

        private void fill_country()
        {
            OleDbCommand cmd = new OleDbCommand("select distinct country from address where country<>''", con);
            OleDbDataReader data;
            if (con.State == ConnectionState.Closed) { con.Open(); }
            data = cmd.ExecuteReader();
            if (data.HasRows == true)
            {
                while (data.Read())
                {
                    country.Add(data["country"].ToString());
                }
            }
            txt_country.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txt_country.AutoCompleteSource = AutoCompleteSource.CustomSource;
            txt_country.AutoCompleteCustomSource = catagry;
            //=-----------------------------------------------------------------------------------------------------------
            OleDbDataAdapter dap = new OleDbDataAdapter("select distinct country from address where country<>''", con);
            DataTable dt = new DataTable();
            dap.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                txt_country.DataSource = dt;
                txt_country.DisplayMember = "country";
            }
        }

        private void fill_Catagory()
        {
            //OleDbDataAdapter dap = new OleDbDataAdapter("select category from address", con);
            OleDbCommand cmd = new OleDbCommand("select distinct category from address where category<>''", con);
            OleDbDataReader data;
            if (con.State == ConnectionState.Closed) { con.Open(); }
            data = cmd.ExecuteReader();
            if (data.HasRows == true)
            {
                while (data.Read())
                {
                    catagry.Add(data["category"].ToString());
                }
            }
            cmd_ctgry.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cmd_ctgry.AutoCompleteSource = AutoCompleteSource.CustomSource;
            cmd_ctgry.AutoCompleteCustomSource = catagry;
            OleDbDataAdapter dap = new OleDbDataAdapter("select distinct category from address where category<>''", con);
            DataTable dt = new DataTable();
            dap.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                cmd_ctgry.DataSource = dt;
                cmd_ctgry.DisplayMember = "category";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
            ID = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txt_name.Text != "")
            {
                string SQL = "";
                if (button1.Text == "&Save")
                {
                    if (dtp_anny.Checked)
                    {
                        SQL = "insert into address(title,company,fname,mno,phoneoffice,phonehome,fax,address1,city,state,country,pin,emailhome,website,category,emailbusiness,profession,spouse,remarks,annvdate,std,serial,centrax) values('" + cmb_title.Text + "','" + txt_companyName.Text + "','" + txt_name.Text + "','" + txt_mob.Text + "','" + txt_phone_offc.Text + "','" + txt_Home_phone.Text + "','" + txt_fax.Text + "','" + txt_address.Text + "','" + txt_city.Text + "','" + txt_State.Text + "','" + txt_country.Text + "','" + txt_Zip.Text + "','" + txt_email_h.Text + "','" + txt_webpage.Text + "','" + cmd_ctgry.Text + "','" + txt_email_B.Text + "','" + txt_profation.Text + "','" + txt_supose.Text + "','" + txt_remarks.Text + "',#" + dtp_anny.Value.ToString("dd-MMM-yy") + "#,'" + txt_STD.Text + "'," + get_seriolNO() + ",'" + txt_contact.Text + "')";
                    }
                    else if (dtp_birth.Checked)
                    {
                        SQL = "insert into address(title,company,fname,mno,phoneoffice,phonehome,fax,address1,city,state,country,pin,emailhome,website,category,emailbusiness,profession,spouse,remarks,bdate,std,serial,centrax) values('" + cmb_title.Text + "','" + txt_companyName.Text + "','" + txt_name.Text + "','" + txt_mob.Text + "','" + txt_phone_offc.Text + "','" + txt_Home_phone.Text + "','" + txt_fax.Text + "','" + txt_address.Text + "','" + txt_city.Text + "','" + txt_State.Text + "','" + txt_country.Text + "','" + txt_Zip.Text + "','" + txt_email_h.Text + "','" + txt_webpage.Text + "','" + cmd_ctgry.Text + "','" + txt_email_B.Text + "','" + txt_profation.Text + "','" + txt_supose.Text + "','" + txt_remarks.Text + "',#" + dtp_birth.Value.ToString("dd-MMM-yy") + "#,'" + txt_STD.Text + "'," + get_seriolNO() + ",'" + txt_contact.Text + "')";
                    }
                    if (!dtp_birth.Checked && !dtp_anny.Checked)
                    {
                        SQL = "insert into address(title,company,fname,mno,phoneoffice,phonehome,fax,address1,city,state,country,pin,emailhome,website,category,emailbusiness,profession,spouse,remarks,std,serial,centrax) values('" + cmb_title.Text + "','" + txt_companyName.Text + "','" + txt_name.Text + "','" + txt_mob.Text + "','" + txt_phone_offc.Text + "','" + txt_Home_phone.Text + "','" + txt_fax.Text + "','" + txt_address.Text + "','" + txt_city.Text + "','" + txt_State.Text + "','" + txt_country.Text + "','" + txt_Zip.Text + "','" + txt_email_h.Text + "','" + txt_webpage.Text + "','" + cmd_ctgry.Text + "','" + txt_email_B.Text + "','" + txt_profation.Text + "','" + txt_supose.Text + "','" + txt_remarks.Text + "','" + txt_STD.Text + "'," + get_seriolNO() + ",'" + txt_contact.Text + "')";
                    }
                    else if (dtp_birth.Checked && dtp_anny.Checked)
                    {
                        SQL = "insert into address(title,company,fname,mno,phoneoffice,phonehome,fax,address1,city,state,country,pin,emailhome,website,category,emailbusiness,profession,spouse,remarks,annvdate,bdate,std,serial,centrax) values('" + cmb_title.Text + "','" + txt_companyName.Text + "','" + txt_name.Text + "','" + txt_mob.Text + "','" + txt_phone_offc.Text + "','" + txt_Home_phone.Text + "','" + txt_fax.Text + "','" + txt_address.Text + "','" + txt_city.Text + "','" + txt_State.Text + "','" + txt_country.Text + "','" + txt_Zip.Text + "','" + txt_email_h.Text + "','" + txt_webpage.Text + "','" + cmd_ctgry.Text + "','" + txt_email_B.Text + "','" + txt_profation.Text + "','" + txt_supose.Text + "','" + txt_remarks.Text + "',#" + dtp_anny.Value.ToString("dd-MMM-yy") + "#,#" + dtp_birth.Value.ToString("dd-MMM-yy") + "#,'" + txt_STD.Text + "'," + get_seriolNO() + ",'" + txt_contact.Text + "')";
                    }
                    OleDbCommand cmd = new OleDbCommand(SQL, con);
                    try
                    {
                        if (con.State == ConnectionState.Closed) { con.Open(); }
                        if (chk())
                        {
                            cmd.ExecuteNonQuery();
                            clear_all();
                            MessageBox.Show("Saved new entry Successful!!", "New Entry", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else { MessageBox.Show("Same name Alreadt exist in the selected Category!", "Already exist", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                    }
                    catch { }
                }
                else
                {
                    try
                    {
                        if (dtp_anny.Checked)
                        {
                            SQL = "update address set title='" + cmb_title.Text + "',company='" + txt_companyName.Text + "',fname='" + txt_name.Text + "',mno='" + txt_mob.Text + "',phoneoffice='" + txt_phone_offc.Text + "',phonehome='" + txt_Home_phone.Text + "',fax='" + txt_fax.Text + "',address1='" + txt_address.Text + "',city='" + txt_city.Text + "',state='" + txt_State.Text + "',country='" + txt_country.Text + "',pin='" + txt_Zip.Text + "',emailhome='" + txt_email_h.Text + "',website='" + txt_webpage.Text + "',category='" + cmd_ctgry.Text + "',emailbusiness='" + txt_email_B.Text + "',profession='" + txt_profation.Text + "',spouse='" + txt_supose.Text + "',remarks='" + txt_remarks.Text + "',annvdate=#" + dtp_anny.Value.ToString("dd-MMM-yy") + "#,std='" + txt_STD.Text + "',centrax='" + txt_contact.Text + "' where id=" + ID;
                        }
                        else if (dtp_birth.Checked)
                        {
                            SQL = "update address set title='" + cmb_title.Text + "',company='" + txt_companyName.Text + "',fname='" + txt_name.Text + "',mno='" + txt_mob.Text + "',phoneoffice='" + txt_phone_offc.Text + "',phonehome='" + txt_Home_phone.Text + "',fax='" + txt_fax.Text + "',address1='" + txt_address.Text + "',city='" + txt_city.Text + "',state='" + txt_State.Text + "',country='" + txt_country.Text + "',pin='" + txt_Zip.Text + "',emailhome='" + txt_email_h.Text + "',website='" + txt_webpage.Text + "',category='" + cmd_ctgry.Text + "',emailbusiness='" + txt_email_B.Text + "',profession='" + txt_profation.Text + "',spouse='" + txt_supose.Text + "',remarks='" + txt_remarks.Text + "',bdate=#" + dtp_birth.Value.ToString("dd-MMM-yy") + "#,std='" + txt_STD.Text + "',centrax='" + txt_contact.Text + "' where id=" + ID;
                        }
                        if (!dtp_birth.Checked && !dtp_anny.Checked)
                        {
                            SQL = "update address set title='" + cmb_title.Text + "',company='" + txt_companyName.Text + "',fname='" + txt_name.Text + "',mno='" + txt_mob.Text + "',phoneoffice='" + txt_phone_offc.Text + "',phonehome='" + txt_Home_phone.Text + "',fax='" + txt_fax.Text + "',address1='" + txt_address.Text + "',city='" + txt_city.Text + "',state='" + txt_State.Text + "',country='" + txt_country.Text + "',pin='" + txt_Zip.Text + "',emailhome='" + txt_email_h.Text + "',website='" + txt_webpage.Text + "',category='" + cmd_ctgry.Text + "',emailbusiness='" + txt_email_B.Text + "',profession='" + txt_profation.Text + "',spouse='" + txt_supose.Text + "',remarks='" + txt_remarks.Text + "',std='" + txt_STD.Text + "',centrax='" + txt_contact.Text + "' where id=" + ID;
                        }
                        else if (dtp_birth.Checked && dtp_anny.Checked)
                        {
                            SQL = "update address set title='" + cmb_title.Text + "',company='" + txt_companyName.Text + "',fname='" + txt_name.Text + "',mno='" + txt_mob.Text + "',phoneoffice='" + txt_phone_offc.Text + "',phonehome='" + txt_Home_phone.Text + "',fax='" + txt_fax.Text + "',address1='" + txt_address.Text + "',city='" + txt_city.Text + "',state='" + txt_State.Text + "',country='" + txt_country.Text + "',pin='" + txt_Zip.Text + "',emailhome='" + txt_email_h.Text + "',website='" + txt_webpage.Text + "',category='" + cmd_ctgry.Text + "',emailbusiness='" + txt_email_B.Text + "',profession='" + txt_profation.Text + "',spouse='" + txt_supose.Text + "',remarks='" + txt_remarks.Text + "',annvdate=#" + dtp_anny.Value.ToString("dd-MMM-yy") + "#,bdate=#" + dtp_birth.Value.ToString("dd-MMM-yy") + "#,std='" + txt_STD.Text + "',centrax='" + txt_contact.Text + "' where id=" + ID;
                        }

                        OleDbCommand cmd = new OleDbCommand(SQL, con);
                        if (con.State == ConnectionState.Closed) { con.Open(); }
                        if (chk1())
                        {
                            cmd.ExecuteNonQuery();
                            ID = 0;
                            clear_all();
                            MessageBox.Show("Updated Successful!!", "Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else { }
                    }
                    catch { }
                }
                this.Close();
            }
            else {
                MessageBox.Show("Name Cann't be Blank!!", "Name Required", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private bool chk1()
        {
            try
            {
                OleDbCommand cmd = new OleDbCommand("select Count(*) from address where fname like '" + txt_name.Text.Trim() + "' and category like '" + cmd_ctgry.Text.Trim() + "'", con);
                if (con.State == ConnectionState.Closed) { con.Open(); }
                if (Convert.ToInt32(cmd.ExecuteScalar()) > 1) { return false; }
                else return true;
            }
            catch { return false; }
        }
        private bool chk()
        {
            try
            {
                OleDbCommand cmd = new OleDbCommand("select Count(*) from address where fname like '" + txt_name.Text.Trim() + "' and category like '" + cmd_ctgry.Text.Trim() + "'", con);
                                if (con.State == ConnectionState.Closed) { con.Open(); }
                                if (Convert.ToInt32(cmd.ExecuteScalar()) > 0) { return false; }
                                else return true;
            }
            catch { return false; }
        }

        private void clear_all()
        {
            //cmb_title.Text = "";
            cmd_ctgry.Text = "";
            txt_name.Clear();
            txt_companyName.Text = "";
            txt_profation.Clear();
            txt_supose.Clear();
            txt_address.Clear();
            txt_city.Text = "";
            txt_State.Text = "";
            txt_STD.Text = "";
            txt_webpage.Clear();
            txt_Zip.Clear();
            txt_contact.Clear();
            txt_country.Text = "";
            txt_email_B.Clear();
            txt_email_h.Clear();
            txt_fax.Clear();
            txt_Home_phone.Clear();
            txt_mob.Clear();
            txt_phone_offc.Clear();
            txt_remarks.Clear();
        }

        private string get_seriolNO()
        {
            try
            {
                OleDbCommand cmd = new OleDbCommand("Select Max(serial) from address", con);
                if (con.State == ConnectionState.Closed) { con.Open(); }
                return cmd.ExecuteScalar().ToString();
            }
            catch { return "1"; }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Are you sure ?" + "\r\n" + "want to delete ?", "You are deleting...", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    OleDbCommand cmd = new OleDbCommand("delete from address where id=" + ID, con);
                    if (con.State == ConnectionState.Closed) { con.Open(); }
                    cmd.ExecuteNonQuery();
                    clear_all();
                    MessageBox.Show("Entry deleted Successfully!!,", "Deleted!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                }
            }
            catch { }
        }

        private void New_Entry_FormClosing(object sender, FormClosingEventArgs e)
        {
            ID = 0;
            code.Clear();
            catagry.Clear();
            country.Clear();
            city.Clear();
            state.Clear();
            cmpny.Clear();
        }

        private void txt_supose_Leave(object sender, EventArgs e)
        {
            txt_address.Focus();
        }

        private void dtp_anny_Leave(object sender, EventArgs e)
        {
            button1.Focus();
        }

              private void saveToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            button1_Click(sender, e);
        }
    }
}
