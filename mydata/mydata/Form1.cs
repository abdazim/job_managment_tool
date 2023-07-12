using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Access; //References-using Excel=Microsoft.Office.Interop.Access;


namespace mydata
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox2.PasswordChar = '*';
            //string path = Directory.GetCurrentDirectory();
            //string filter = @"\my.accdb";
            //string fullpath = path + filter;
            //MessageBox.Show(fullpath);

            //all offices from 2007-2016 contain the provider "Microsoft.ACE.Oledb.12.0"
            //Microsoft Access Database Engine 2010 : 
            //https://www.microsoft.com/en-us/download/details.aspx?id=13255 
        }



        /// <summary>
        /// Log file create function
        /// </summary>
        /// <param name="sEventName"></param>
        /// <param name="sControlName"></param>
        /// <param name="sFormName"></param>
        public void LogFile(string sEventName, string sControlName, string sFormName)
        {
            StreamWriter log;
            if (!File.Exists("logfile.txt"))
            {
                log = new StreamWriter("logfile.txt");
            }
            else
            {
                log = File.AppendText("logfile.txt");
            }
            // Write to the file:
            log.WriteLine("===============================================Srart============================================");
            log.WriteLine("Data Time:" + DateTime.Now);
            log.WriteLine("--------------");
            //log.WriteLine("Exception Name:" + sExceptionName);
            log.WriteLine("Event Name:" + sEventName);
            log.WriteLine("---------------");
            log.WriteLine("Control Name:" + sControlName);
            log.WriteLine("---------------");
            log.WriteLine("Form Name:" + sFormName);
            log.WriteLine("===============================================End==============================================");
            // Close the stream:
            log.Close();
        }

        /// <summary>
        /// check username and password in the database 
        /// </summary>
        /// <returns></returns>
        public bool CheckFunction()
        {
           var conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my.accdb;");
           conn.Open();
            try
            {
                DataTable dt = new DataTable();
                DataSet ds = new DataSet();
                ds.Tables.Add(dt);
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT user, pass FROM users;", conn);
                da.Fill(dt);
                foreach (DataRow r in dt.Rows)
                {
                    if (r[0].ToString() == textBox1.Text && r[1].ToString() == textBox2.Text)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }            
            }
            catch (Exception ex)
            {
                LogFile("login -CheckFunction-Exception " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }
            finally { conn.Close(); }
            return false;
        }


        /// <summary>
        ///  login  - Enter button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (CheckFunction() == true)
                {
                    panel1.Visible = true;
                    panel1.Dock = DockStyle.Fill;
                }
                else
                {
                    string message = "Sorry, your login information is not correct. Please try again.";
                    MessageBox.Show(message);
                    LogFile(sender + " - login information Eror Message: " + message+ "- username or password wrong: "+"user: "+textBox1.Text+","+"pass: "+ textBox2.Text, ((Control)sender).Name, this.FindForm().Name);
                }
            }  
	        catch (Exception ex)
	        {
                LogFile("login - Enter button-Exception " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }
        }


        /// <summary>
        /// exit
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        /// <summary>
        /// ToolStrip - about 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void aboutToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }



        ///panel1///////////////////////////////////////////////////////////////
        ///
        /// <summary>
        /// panel1 users button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
        }


        /// <summary>
        /// panel1 products button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            product product = new product();
            product.Show();
        }


        /// <summary>
        /// panel1 Customers button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            //
        }


        private void button6_Click(object sender, EventArgs e)
        {
            var About = new About();
            About.Show();
        }


        ////panel1 ToolStrip////////////////////////////////////////////

        /// <summary>
        /// panel1 instuctions 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void instuctionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var Instructions = new Instructions();
            Instructions.Show();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }





        ///////////////////////////////////////////////////////////////////////////////////////////////

    }
}
