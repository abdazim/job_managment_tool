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
using System.Reflection;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Data.SqlClient;

namespace mydata
{
    public partial class Form2 : Form
    {
        //https://sourceforge.net/projects/itextsharp/?source=typ_redirect
        static string constring=@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my.accdb";
        OleDbConnection con = new OleDbConnection(constring);
        OleDbCommand cmd;
        OleDbDataAdapter adapter;
        DataTable dt = new DataTable();
        public Form2()  //constructor 
        {
            InitializeComponent();
            //textBox2.PasswordChar = '*';
            BindDataGridView();
        }

        // to access to form1 functions ... example - form1.LogFile(ex.Message, ex.StackTrace, this.FindForm().Name);
        Form1 form1 = new Form1();    


        /// <summary>
        /// dataGridView , Table Header create 
        /// </summary>
        private void BindDataGridView()
        {
            //creat header of the datagridview show
            dataGridView1.ColumnCount = 3;
            dataGridView1.Columns[0].Name = "ID";
            dataGridView1.Columns[1].Name = "user";
            dataGridView1.Columns[2].Name = "pass";
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            //dataGridView Properties
            //selection mode
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = false;
            dataGridView1.AllowUserToAddRows = false; //disable the  empty row at the bottom of a DataGridView. It allows the user to add new data at run-time
            dataGridView1.Columns[0].HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9, FontStyle.Bold);
            dataGridView1.Columns[1].HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9, FontStyle.Bold);
            dataGridView1.Columns[2].HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9, FontStyle.Bold);

            //show data (users table) auto show without button
            retrieve();
        }

        /////
        //check if user exist in the database 
        /// <summary>
        /// check exist function
        /// </summary>
        /// <param name="user"></param>
        /// <param name="pass"></param>     
        private void exist(string user,string pass)
        {
            try
            { 
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandText = "SELECT [user] FROM users WHERE user = @user";
                cmd.Parameters.AddWithValue("@user", user);
                con.Open();
                //the EROR - The connection was not closed. The connection's current state is Open mean you have conn.Open() 2 times in your method.should remove one.
                OleDbDataReader rdr = cmd.ExecuteReader();

                if (rdr.HasRows == true)
                {
                    MessageBox.Show("The user is Exist");
                    con.Close();
                }
                else
                {
                    add(textBox1.Text, textBox2.Text);
                }
            con.Close();         
            retrieve();  //REFRESH
            }  
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();             
                form1.LogFile(" check exist function-Eror Message: " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }
        }



        //Add////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// Add user function
        /// </summary>
        /// <param name="user"></param>
        /// <param name="pass"></param>     
        private void add(string user,string pass)
        {
            //sql
            string sql = "INSERT into users ([user],[pass]) VALUES (@user,@pass)";
            cmd = new OleDbCommand(sql,con);

            //add parm
            cmd.Parameters.AddWithValue("@user", user);
            cmd.Parameters.AddWithValue("@pass", pass);

            //opem con exec
            try
            {
                if (cmd.ExecuteNonQuery() > 0)
                {
                    MessageBox.Show("successfuly inserted");
                }
                con.Close();
                //REFRESH
                retrieve();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();             
                form1.LogFile("  Add user function-Eror Message: " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }
        }


        /// <summary>
        /// Add button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(textBox1.Text) || String.IsNullOrEmpty(textBox2.Text))
            {
                string message="Please enter your user name / password ";
                MessageBox.Show(message);
                form1.LogFile(sender + " - Add button-Eror Message: " + message + "- username or password empty: " + "user: " + textBox1.Text + "," + "pass: " + textBox2.Text, ((Control)sender).Name, this.FindForm().Name);
            }
            else
            {
                 exist(textBox1.Text, textBox2.Text); //check user exist
            }                 
        }


        ///textbox//////////////////////////////////////////////////////////////////
        /// <summary>
        /// user textbox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox1.Text, "[^A-Za-z0-9]"))
            {
                ///place cursor at the end,save position textbox at the end after remove length-1
                bool focused = textBox1.Focused;
                int start = textBox1.SelectionStart;
                int len = textBox1.SelectionLength;
                ///remove the wrong digit
                string message = "Enter a valid User Name";
                MessageBox.Show(message);
                textBox1.Text = textBox1.Text.Remove(textBox1.Text.Length - 1);
                ///restore to the last position cursor
                textBox1.SelectionStart = start;
                textBox1.SelectionLength = len;
                textBox1.Select();
                form1.LogFile(sender + " - textBox1 - Eror Message: " + message + "-form2 Textbox1: " + "user: " + textBox1.Text, ((Control)sender).Name, this.FindForm().Name);
            }

        }

        /// <summary>
        /// password textbox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox2.Text, "[^0-9a-zA-Z]"))
            {
                ///place cursor at the end,save position textbox at the end after remove length-1
                bool focused = textBox2.Focused;
                int start = textBox2.SelectionStart;
                int len = textBox2.SelectionLength;
                ///remove the wrong digit
                string message = "Enter a valid User Name";
                MessageBox.Show(message);
                textBox2.Text = textBox2.Text.Remove(textBox2.Text.Length - 1);
                ///restore to the last position cursor
                textBox2.SelectionStart = start;
                textBox2.SelectionLength = len;
                textBox2.Select();
                form1.LogFile(sender + " - textBox2 - Eror Message: " + message + "-form2 Textbox2: " + "pass: " + textBox2.Text, ((Control)sender).Name, this.FindForm().Name);
            }
        }



        //update////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// update user function 
        /// </summary>
        /// <param name="id"></param>
        /// <param name="user"></param>
        /// <param name="pass"></param>
        private void update(int id, string user, string pass)
        {
            //sql stmt
            string sql = "UPDATE users SET [pass]='" + pass + "' WHERE [user]= '" + user + "'";  //value will be with ' ' ;
            cmd = new OleDbCommand(sql, con);
            //open com,update,retrieve dgview
            try
            {
                con.Open();
                adapter = new OleDbDataAdapter(cmd);
                adapter.UpdateCommand = con.CreateCommand();
                adapter.UpdateCommand.CommandText = sql;
                if (adapter.UpdateCommand.ExecuteNonQuery() > 0)
                {
                    clearTxts();
                    MessageBox.Show("successfuly updated");
                }
                else
                { MessageBox.Show("Enter a correct user for update "); }
                con.Close();
                retrieve();  //REFRESH
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                form1.LogFile("update user function-Exception - " + ex.Message, ex.StackTrace, this.FindForm().Name);
                con.Close();
            }
        }


        /// <summary>
        /// update button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            String selected = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
            int id = Convert.ToInt32(selected);
            update(id, textBox1.Text, textBox2.Text); //function
        }


        //Delete//////////////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// DELETE user function
        /// </summary>
        /// <param name="id"></param>
        private void delete(int id)
        {
            ////SQL STMT
            String sql = "DELETE FROM users WHERE [ID]=" + id + "";
            cmd = new OleDbCommand(sql, con);
            //'OPEN CON,EXECUTE DELETE,CLOSE CON
            try
            {
                con.Open();
                adapter = new OleDbDataAdapter(cmd);
                adapter.DeleteCommand = con.CreateCommand();
                adapter.DeleteCommand.CommandText = sql;
                //PROMPT FOR CONFIRMATION
                if (MessageBox.Show("Are You Sure? ", "DELETE", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    if (cmd.ExecuteNonQuery() > 0)
                    {
                        MessageBox.Show("Successfully deleted");
                    }
                }
                con.Close();
                retrieve();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                form1.LogFile("DELETE user function -Exception - " + ex.Message, ex.StackTrace, this.FindForm().Name);
                con.Close();
            }
        }


        /// <summary>
        /// delete button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            String selected = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
            int id = Convert.ToInt32(selected);
            delete(id);
        }




        //PDF//////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// PDF button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            //Creating iTextSharp Table from the DataTable data
            PdfPTable pdfTable = new PdfPTable(dataGridView1.ColumnCount);
            pdfTable.DefaultCell.Padding = 3;
            pdfTable.WidthPercentage = 30;
            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
            pdfTable.DefaultCell.BorderWidth = 1;

            //Adding Header row
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                cell.BackgroundColor = new iTextSharp.text.Color(240, 240, 240);
                pdfTable.AddCell(cell);
            }

            //Adding DataRow
            try
            {   
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.Value != null)
                        { 
                            pdfTable.AddCell(cell.Value.ToString());
                        }
                    }
                }              
            }
             catch (Exception ex)
            {
                MessageBox.Show(ex.Message); ;
                form1.LogFile("PDF button-Exception-  " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }

            //Exporting to PDF
            //string folderPath = "C:\\PDFs\\";
            string folderPath = Directory.GetCurrentDirectory() + @"\PDFs\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            using (FileStream stream = new FileStream(folderPath + "DataGridViewExport.pdf", FileMode.Create))
            {
                Document pdfDoc = new Document(PageSize.A2, 10f, 10f, 10f, 0f);
                PdfWriter.GetInstance(pdfDoc, stream);
                pdfDoc.Open();
                pdfDoc.Add(pdfTable);
                pdfDoc.Close();
                stream.Close();
            }
        }



        //FILL data /////////////////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// FILL Datagridview function
        /// </summary>
        /// <param name="id"></param>
        /// <param name="user"></param>
        /// <param name="pass"></param>
        private void filldgv(string id, string user, string pass)
        {
            dataGridView1.Rows.Add(id, user, pass);
        }


        /// <summary>
        /// retrival(shelfat ntonem) of data function
        /// </summary>
        private void retrieve()
        {
            dataGridView1.Rows.Clear();
            string sql = "SELECT * FROM users;";
            cmd = new OleDbCommand(sql, con);
            try
            {
                con.Open();
                adapter = new OleDbDataAdapter(cmd);
                adapter.Fill(dt);
                //loop thru dt
                foreach (DataRow row in dt.Rows)
                { filldgv(row[0].ToString(), row[1].ToString(), row[2].ToString()); }
                con.Close();
                //clear dt 
                dt.Rows.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message); ;
                form1.LogFile("retrival-shelfat ntonemof data function- Exception - " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }
        }


        //clear and exit//////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// clear txt function
        /// </summary>
        private void clearTxts()
        {
            textBox1.Text = "";
            textBox2.Text = "";
        }


        /// <summary>
        /// clear button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            //dataGridView1.Rows.Clear();
            clearTxts();
        }


        /// <summary>
        /// exit
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }


        ///////////////////////////////////////////////////////////////////////////////////////////////


    }
}
