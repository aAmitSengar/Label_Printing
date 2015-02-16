using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Label
{
    class access_data
    {
    }
    public sealed class Dataaccess
    {
        // private static string Path = System.IO.Directory.GetParent(Application.StartupPath).ToString();
        public static readonly string connection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\AddressBook.accdb;";

    }
    public class cc2
    {
        public void newCommand_ExecuteNonQuery(string sql, OleDbConnection ccn)
        {
            try
            {
                if (ccn.State == ConnectionState.Closed) { ccn.Open(); }
                OleDbCommand cmd = new OleDbCommand(sql, ccn);

                cmd.ExecuteNonQuery();
            }
            catch { }
        }
        public double newCommand_ExecuteScaler(string sql, OleDbConnection ccn)
        {
            try
            {
                if (ccn.State == ConnectionState.Closed) { ccn.Open(); }
                OleDbCommand cmd = new OleDbCommand(sql, ccn);
                return Convert.ToDouble(cmd.ExecuteScalar());
            }
            catch { return 0; }
        }
        public string newCommand_ExecuteScaler_string(string sql, OleDbConnection ccn)
        {
            try
            {
                if (ccn.State == ConnectionState.Closed) { ccn.Open(); }
                OleDbCommand cmd = new OleDbCommand(sql, ccn);
                return cmd.ExecuteScalar().ToString();
            }
            catch { return ""; }
        }
        public DataTable newCommand_Dataadapter(string sql, OleDbConnection ccn)
        {
            DataTable dt = new DataTable();
            try
            {
                if (ccn.State == ConnectionState.Closed) { ccn.Open(); }
                OleDbDataAdapter cmd = new OleDbDataAdapter(sql, ccn);
                cmd.Fill(dt);
                return dt;
            }
            catch { return dt; }
        }
        public void newCommand_DELETE(string sql, OleDbConnection ccn)
        {
            try
            {
                if (ccn.State == ConnectionState.Closed) { ccn.Open(); }
                OleDbCommand cmd = new OleDbCommand(sql, ccn);
                cmd.ExecuteNonQuery();
            }
            catch { }
        }

        public bool newCommand_ExecuteScaler_bool(string SQL, OleDbConnection con)
        {
            try
            {
                if (con.State == ConnectionState.Closed) { con.Open(); }
                OleDbCommand cmd = new OleDbCommand(SQL, con);
                return Convert.ToBoolean(cmd.ExecuteScalar().ToString());
            }
            catch { return false; }
        }
    }
}
