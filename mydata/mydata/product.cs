using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Access; //References-using Excel=Microsoft.Office.Interop.Access;
using System.Reflection;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Data.OleDb;
using mydata.Properties;
using System.Drawing.Imaging;  //use System.Drawing.Color.White; in update product button 
using System.Threading;

namespace mydata
{
    public partial class product : Form
    {
        //https://sourceforge.net/projects/itextsharp/?source=typ_redirect
        static string constring = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=my.accdb";
        OleDbConnection con = new OleDbConnection(constring);
        OleDbCommand cmd;
        OleDbDataAdapter adapter;
        DataTable dt = new DataTable();
        //int imageIDX = 1;       

        public product()
        {
            InitializeComponent();
            BindDataGridView();
        }

        // to access to form1 functions ... example - form1.LogFile(ex.Message, ex.StackTrace, this.FindForm().Name);
        Form1 form1 = new Form1();



        /// <summary>
        /// dataGridView , Table header create
        /// </summary>
        private void BindDataGridView()
        {
            //creat header of the datagridview show
            dataGridView1.ColumnCount = 10;
            dataGridView1.Columns[0].Name = "ID";
            dataGridView1.Columns[1].Name = "name";
            dataGridView1.Columns[2].Name = "color";
            dataGridView1.Columns[3].Name = "size";
            dataGridView1.Columns[4].Name = "weight";
            dataGridView1.Columns[5].Name = "count";
            dataGridView1.Columns[6].Name = "price";
            dataGridView1.Columns[7].Name = "Description";
            dataGridView1.Columns[8].Name = "Date Time";
            dataGridView1.Columns[9].Name = "Image";

            //dataGridView Properties
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //selection mode
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = false;
            dataGridView1.AllowUserToAddRows = false; //disable the  empty row at the bottom of a DataGridView. It allows the user to add new data at run-time

            //header style
            dataGridView1.Columns[0].HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9, FontStyle.Bold);
            dataGridView1.Columns[1].HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9, FontStyle.Bold);
            dataGridView1.Columns[2].HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9, FontStyle.Bold);
            dataGridView1.Columns[3].HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9, FontStyle.Bold);
            dataGridView1.Columns[4].HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9, FontStyle.Bold);
            dataGridView1.Columns[5].HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9, FontStyle.Bold);
            dataGridView1.Columns[6].HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9, FontStyle.Bold);
            dataGridView1.Columns[7].HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9, FontStyle.Bold);
            dataGridView1.Columns[8].HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9, FontStyle.Bold);
            dataGridView1.Columns[9].HeaderCell.Style.Font = new System.Drawing.Font("Tahoma", 9, FontStyle.Bold);
         
            //show data (users table) auto show without button
            retrieve();
        }


        //Add////////////////////////////////////////////////////////////////////////////////////

        /// <summary>
        /// Add product function
        /// </summary>
        /// <param name="user"></param>
        /// <param name="pass"></param>    
        /// 

        private void add(string name, string color,string size,double weight,Int32 count,double price,string Description,string Date_Time ,string image1)
        {
            //sql
            //MessageBox.Show(image1);
            string sql = "INSERT into products ([name],[color],[size],[weight],[count],[price],[Description],[Date_Time], [image1]) VALUES (@name,@color,@size,@weight,@count,@price,@Description,@Date_Time,@image1)";
            cmd = new OleDbCommand(sql, con);

            //add parm
            cmd.Parameters.AddWithValue("@name", name);
            cmd.Parameters.AddWithValue("@color", color);
            cmd.Parameters.AddWithValue("@size", size);
            cmd.Parameters.AddWithValue("@weight", weight);
            cmd.Parameters.AddWithValue("@count", count);
            cmd.Parameters.AddWithValue("@price", price);
            cmd.Parameters.AddWithValue("@Description", Description);
            cmd.Parameters.AddWithValue("@Date_Time", Date_Time);
            cmd.Parameters.AddWithValue("@image1", image1);


            //open con exec
            try
            {
                con.Open();
                if (cmd.ExecuteNonQuery() > 0)
                {
                    MessageBox.Show("successfuly inserted");
                }
                con.Close();
                retrieve(); //REFRESH
                              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
                form1.LogFile("Add product function-Eror Message: " + ex.Message, ex.StackTrace, this.FindForm().Name);

            }
        }



        /// <summary>
        /// load image button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        string nopicload = @".\no.jpg";
        private void button7_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = null;
            label9.Text = null;

            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            string str;
            //image filter
            openFileDialog1.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    str = openFileDialog1.FileName;
                    textBox8.Text = str.ToString();
                    //MessageBox.Show(str);
                    pictureBox1.Image = new Bitmap(openFileDialog1.FileName); // display image in picture box 

                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {

                        }
                    }
                    //MessageBox.Show("load image button: "+str.ToString());
                }

                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                    form1.LogFile("load image function Eror Message: " + ex.Message, ex.StackTrace, this.FindForm().Name);
                }
            }

            else // don't select picture window , put the nopicload (no.jpg)
            {
                if (File.Exists(nopicload))
                {
                    textBox8.Text = nopicload;
                    System.Drawing.Image img  ;
                    img = System.Drawing.Image.FromFile(nopicload);
                    pictureBox1.Image = img;                 
                }
                else
                { MessageBox.Show("Can't Find The (NO.JPG) Picture In The Base Directory."); } // if can't find the pic
            }

        }


        /// <summary>
        /// add product button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        string imagePath;
        string timenow = DateTime.Now.ToLocalTime().ToString(); //SENd to (add) func 
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (String.IsNullOrEmpty(textBox1.Text) || String.IsNullOrEmpty(textBox2.Text) || String.IsNullOrEmpty(textBox3.Text) || String.IsNullOrEmpty(textBox4.Text) || String.IsNullOrEmpty(textBox5.Text) || String.IsNullOrEmpty(textBox6.Text) || String.IsNullOrEmpty(textBox8.Text))
                {
                    MessageBox.Show("Sorry, information required, please fill  ");
                }
                else
                {
                    pic_check_and_save_to_pc(); // check if picturebox/textbox8 appear 
                    checkvalue(); // check textbox4+textbox5+testbox6                                                
                    clearTxts();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("add_product_button:  " + ex.Message);
                form1.LogFile("add product button : " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }
        }



        /// <summary>
        /// row id 
        /// </summary>
        /// 
        string row_id_data;
        private string get_row_id()
        {
            row_id_data = dataGridView1.CurrentRow.Cells[0].Value.ToString(); //id of row
            return row_id_data;
            //////////////////////////////////////////////
            //http://webmaster.org.il/articles/csharp-boolean-operators
            // && -  if  if both true -> return true  ,  true && true  -> return true 
            // || -  if  if one of them true -> return true 
            //////////////////////////////////////////////
        }



        /// <summary>
        /// pictureBox1.Image check / textBox8.Text check if not null
        /// </summary>
        /// 
        private void pic_check_and_save_to_pc()
        {
            string FilePath;
            try
            {
                //check if images folder exist in base directory
                if (Directory.Exists(Application.StartupPath + "\\Images") != true)
                { // if not exsist create folder 
                    Directory.CreateDirectory(Application.StartupPath + "\\Images");
                }
                //check row id ///////////////////////
                if (get_row_id() != null)
                {
                    if (pictureBox1.Image != null)
                        if (textBox8.Text != null)
                        {
                            string newdatesave = DateTime.Now.ToString("dd_MM_yyyy_hh_mm_ss");
                            //FilePath = Directory.GetCurrentDirectory() + "\\Images\\" + newdatesave + ".jpg";
                            FilePath = Application.StartupPath + "\\Images\\" + ("ID_" + get_row_id() + "_DT_" + newdatesave) + ".jpg";
                            pictureBox1.Image.Save(FilePath, ImageFormat.Jpeg);
                            imagePath = FilePath;
                        }
                        else //if pictureBox1.Image ==null or textBox8.Text ==null
                        {
                            // imagePath = null;
                            System.Drawing.Image img1;
                            using (img1 = System.Drawing.Image.FromFile(nopic))
                            {
                                pictureBox1.Image = img1;
                            }
                            img1.Dispose();
                        }

                }
                else { MessageBox.Show("you have no products"); }
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show("pic_check_and_save_to_pc()-FileNotFoundException" + ex.Message);
            }
            catch (IOException ioEx)
            {
                MessageBox.Show("pic_check_and_save_to_pc()-IOException-Another user is already using this file." + ioEx.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("pic_check_and_save_to_pc():  " + ex.Message);
                form1.LogFile("pic_check_and_save_to_pc() : " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }

        }



        /// <summary>
        /// checkvalues of textboxes before add product
        /// </summary>
        /// 
        private void checkvalue()
        {
            //Int16 -- (-32,768 to +32,767) , Int32 -- (-2,147,483,648 to +2,147,483,647) , Int64 -- (-9,223,372,036,854,775,808 to +9,223,372,036,854,775,807)
            //double 15 - 16 digits
            //check id textbox1,2,3,4,5,6 is not empty
            //if not empty send this data to function , imagepath=when select a pic get the link
            double dValue4; //textbox4 value (mshkal) shoud be double 
            dValue4 = Convert.ToDouble(textBox4.Text); //convert textbox4 to double to check numbers 

            Int32 ivalue; //textbox5 (kmot) should be int
            ivalue = Convert.ToInt32(textBox5.Text);//convert textbox5 to int32 

            double dValue;  //textbox6 value (m7er) shoud be double 
            dValue = Convert.ToDouble(textBox6.Text); //convert textbox to double to check numbers 

            //dValue4=textbox4 , ivalue= textbox5 , dValue=textbox6 , 
            try
            {
                if (dValue4 > 0 && dValue4 <= 50000) //check if textbox4 (mshkal) between 0 and 50,000
                {
                    if (ivalue > 0 && ivalue <= 100000) //check if textbox5 (kmot) between 0 and 100,000
                    {
                        if (dValue > 0 && dValue <= 8000000) //check if textbox6 (price) between 0 and 8,000,000
                        {
                            add(textBox1.Text, textBox2.Text, textBox3.Text, Convert.ToDouble(textBox4.Text), Convert.ToInt32(textBox5.Text), Convert.ToDouble(textBox6.Text), textBox7.Text, timenow, imagePath);
                        }
                        else
                        {
                            pictureBox2.Visible = true;
                        }
                    }
                    else
                    {
                        pictureBox3.Visible = true;
                    }
                }
                else
                {
                    pictureBox4.Visible = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("checkvalue():  " + ex.Message);
                form1.LogFile("checkvalue() : " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }

            //hide the x picture after button press 
            pictureBox2.Visible = false;
            pictureBox3.Visible = false;
            pictureBox4.Visible = false;
        }



        /// <summary>
        /// dataGridView Cell Mouse Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        ///  
        string nopic; // also used to add button 
        string image_grid_value;    //=val = datagridview image path  
        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                // get image path when press on datagridview row 
                DataGridViewSelectedRowCollection rows = dataGridView1.SelectedRows;
                string val = (string)rows[0].Cells[@"image"].Value; //show image path (row0) from datagridview
                image_grid_value = val;


                //get_row_id when press on the row data grid view - shows in label 9  ;
                string row_id_data_Grid = dataGridView1.CurrentRow.Cells[0].Value.ToString(); //id of row
      
                //get Empty pic from directory 
                nopic = @".\no.jpg";

                //check if (val)-image path exist in the folder and not null string 
                if (File.Exists(val) && val != null)
                {
                    System.Drawing.Image img;
                    img = System.Drawing.Image.FromFile(val);
                    pictureBox1.Image = img;   //show the relevant pic in picturebox1
                    label9.Visible = true;
                    label9.Text = row_id_data_Grid.ToString();
                }
                else 
                {
                    System.Drawing.Image img;
                    img = System.Drawing.Image.FromFile(nopic);
                    pictureBox1.Image = img;   //show the pic in picturebox1
                    label9.Visible = true;
                    label9.Text = row_id_data_Grid.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("dataGridViewCell Mouse Click- " + ex.Message);
                form1.LogFile("dataGridViewCell Mouse Click: " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }
        }


        //update////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// update product function 
        /// </summary>
        /// <param name="id"></param>
        /// <param name="user"></param>
        /// <param name="pass"></param>
        private void update(int id, string name, string color , string size , double weight , int count, double price, string Description, string Date_Time, string image1)
        {
            //, string size, double weight, Int32 count, double price, string Description, string Date_Time, string image1
            //int id,string name, string color, string size, double weight, int count, double price, string Description, string Date_Time, string image1)

            string sql1 = "UPDATE [products] SET [name]='" + name + "', [color]='"+ color + "', [size]='" + size + "', [weight]='" + weight + "', [count]='" + count + "', [price]='" + price + "', [Description]='" + Description + "', [Date_Time]='" + Date_Time + "', [image1]='" + image1 + "' WHERE [ID]= " + id + "";  //value will be with ' ' ;
            cmd = new OleDbCommand(sql1, con);
            try
            {
                con.Open();
                adapter = new OleDbDataAdapter(cmd);
                adapter.UpdateCommand = con.CreateCommand();
                adapter.UpdateCommand.CommandText = sql1;
                //PROMPT FOR CONFIRMATION
                if (MessageBox.Show("Are You Sure? ", "Update", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {            
                    if (adapter.UpdateCommand.ExecuteNonQuery() > 0)
                    {                   
                        MessageBox.Show("Product Id: "+ id + ", successfuly updated");
                    }
                    else
                    { MessageBox.Show("Enter a correct Product for update "); }
                }
                con.Close();
                retrieve();  //REFRESH
            }
            catch (Exception ex)
            {
                MessageBox.Show("update product-" + ex.Message);
                form1.LogFile("update product function : " + ex.Message, ex.StackTrace, this.FindForm().Name);
                con.Close();
            }
        }


        /// <summary>
        /// check datagridview row[i]->(shorot)- cells[0]->id compare to get_row_id()->current row id if selected 
        /// </summary>
        string datagridview_id;
        string datagridview_name;
        string datagridview_color;
        string datagridview_size;
        string datagridview_weight;
        string datagridview_count;
        string datagridview_price;
        string datagridview_image;
        private void checkdatagridview()
        {
            try
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    //check datagridview row[i]->(shorot)- cells[0]->id compare to get_row_id()->current row id if selected
                    datagridview_id = (dataGridView1.Rows[i].Cells[0].Value as string);
                    if (get_row_id() == datagridview_id)
                    { // get all values from selected row and put to variables 
                        datagridview_name = (dataGridView1.Rows[i].Cells[1].Value as string);
                        datagridview_color = (dataGridView1.Rows[i].Cells[2].Value as string);
                        datagridview_size = (dataGridView1.Rows[i].Cells[3].Value as string);
                        datagridview_weight = (dataGridView1.Rows[i].Cells[4].Value.ToString());
                        datagridview_count = (dataGridView1.Rows[i].Cells[5].Value.ToString());
                        datagridview_price = (dataGridView1.Rows[i].Cells[6].Value.ToString());
                        datagridview_image = (dataGridView1.Rows[i].Cells[9].Value.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("checkdatagridview() -" + ex.Message);
                form1.LogFile("checkdatagridview() : " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }
        }



        /// <summary>
        /// update product button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void button3_Click(object sender, EventArgs e)
        {
            int id = Convert.ToInt16(get_row_id()); //convert returned number function get_row_id to int16

            //DataGridViewSelectedRowCollection rows2 = dataGridView1.SelectedRows;
            //string val_update = (string)rows2[0].Cells[@"image"].Value; //image path from datagridview(image)        
            //rows2[0].Cells[@"image"].Value= null;
            checkdatagridview(); //check textboxes before use ypdate func 

            string val_update = datagridview_image;

            try
            {
                //if (String.IsNullOrEmpty(textBox1.Text) && String.IsNullOrEmpty(textBox2.Text) && String.IsNullOrEmpty(textBox3.Text) && String.IsNullOrEmpty(textBox4.Text) && String.IsNullOrEmpty(textBox5.Text) && String.IsNullOrEmpty(textBox6.Text))      
                //if some of values empty enter to func to check witch values empty 
                if (String.IsNullOrEmpty(textBox1.Text) || String.IsNullOrEmpty(textBox2.Text) || String.IsNullOrEmpty(textBox3.Text) || String.IsNullOrEmpty(textBox4.Text) || String.IsNullOrEmpty(textBox5.Text) || String.IsNullOrEmpty(textBox6.Text))                     
                    {  
                    if (String.IsNullOrEmpty(textBox1.Text) && String.IsNullOrEmpty(textBox2.Text) && String.IsNullOrEmpty(textBox3.Text) && String.IsNullOrEmpty(textBox4.Text) && String.IsNullOrEmpty(textBox5.Text) && String.IsNullOrEmpty(textBox6.Text))
                    { MessageBox.Show("No Values Entered ,Please Enter Values "); return; } // if no values entred shpw message and exit(return)

                    //after copy the original value if the textbox empty you see the original value in text box (to hide this value use //textcolor=white ).
                    if (String.IsNullOrEmpty(textBox1.Text)) { textBox1.ForeColor = System.Drawing.Color.White; textBox1.Text = datagridview_name; } //color use function System.Drawing
                    if (String.IsNullOrEmpty(textBox2.Text)) { textBox2.ForeColor = System.Drawing.Color.White; textBox2.Text = datagridview_color; }
                    if (String.IsNullOrEmpty(textBox3.Text)) { textBox3.ForeColor = System.Drawing.Color.White; textBox3.Text = datagridview_size; }
                    if (String.IsNullOrEmpty(textBox4.Text)) { textBox4.ForeColor = System.Drawing.Color.White; textBox4.Text = datagridview_weight; }
                    if (String.IsNullOrEmpty(textBox5.Text)) { textBox5.ForeColor = System.Drawing.Color.White; textBox5.Text = datagridview_count; }
                    if (String.IsNullOrEmpty(textBox6.Text)) { textBox6.ForeColor = System.Drawing.Color.White; textBox6.Text = datagridview_price; }
                    }
                //
                /////////////////////////////////////////////////////////////////////////////
                //http://webmaster.org.il/articles/csharp-boolean-operators
                // && -  if  if both true -> return true  ,  true && true  -> return true 
                // || -  if  if one of them true -> return true 
                /////////////////////////////////////////////////////////////////////////////

                if (get_row_id() != null) //if select row != null
                {
                    if (String.IsNullOrEmpty(textBox8.Text) || textBox8.Text == @".\no.jpg") //if textbox 8 null or no photo
                    {
                        imagePath = val_update; //save the current imagepath from database ((string)rows2[0].Cells[@"image"].Value;)
                         
                    }
                    else //if picture is choosen enter to func save pic to pc 
                    {
                        pic_check_and_save_to_pc(); // check if picturebox/textbox8 appear  
                    }

                    update(id, textBox1.Text, textBox2.Text, textBox3.Text, Convert.ToDouble(textBox4.Text), Convert.ToInt16(textBox5.Text), Convert.ToDouble(textBox6.Text), textBox7.Text, timenow, imagePath);
                    clearTxts(); // clear all textbox and pic after updated the values 
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("update button -" + ex.Message);
                form1.LogFile("update product button : " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }
        }



        //Delete//////////////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// DELETE row product function
        /// </summary>
        /// <param name="id"></param>
        private void deleterow(int id)
        {
            
            ////SQL STMT
            String sql = "DELETE FROM products WHERE [ID]=" + id + ""; // delete all row from sql when id = id 
            //string picdel = Application.StartupPath + "\\Images\\" + +id+".jpg";
            //
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
                MessageBox.Show("deleterow function-" + ex.Message);
                form1.LogFile("deleterow function : " + ex.Message, ex.StackTrace, this.FindForm().Name);
                con.Close();
            }
        }



        /// <summary>
        /// deletefile
        /// </summary>
        /// 
        private void deletefile(string image_val)
        {
            try
            {
                ////MessageBox.Show("deletefilefunc: " + val11);
                if (File.Exists(image_val))
                {

                    File.Delete(image_val);
                    MessageBox.Show("Successfully deleted");
                }
                else { MessageBox.Show("File Not Found"); }

            }

            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show("deletefile-UnauthorizedAccessException- " + ex.Message);
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show("deletefile-ArgumentNullExceptione" + ex.Message);
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show("deletefile-FileNotFoundException" + ex.Message);
            }
            catch (IOException ioEx)
            {
                MessageBox.Show("deletefile-IOException-Another user is already using this file." + ioEx.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("deletefile-" + ex.Message);
                form1.LogFile("deletefile : " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private string image_path()
        {
            //string image_val;
            //image_val = null;
            //DataGridViewSelectedRowCollection rows = dataGridView1.SelectedRows;
            //image_val = (string)rows[0].Cells[@"image"].Value; //image path from datagridview(image)
            //return image_val;
            checkdatagridview();
            string val_delete = datagridview_image;
            return val_delete;

        }


        /// <summary>
        /// delete product button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void button4_Click(object sender, EventArgs e)
        {
          string row_id = dataGridView1.CurrentRow.Cells[0].Value.ToString(); //row id (row selected)
          int int_row_id =  Convert.ToInt16(row_id); // for delete func 
            try
            {
                //check if images folder exist in base directory
                if (!Directory.Exists(Application.StartupPath + "\\Images") != true) // if not exsist   
                {           
                    if (row_id != null)
                    {
                       //MessageBox.Show(int_row_id.ToString());
                       //deleterow(int_row_id);
                       deletefile(image_path());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("DELETE button-" + ex.Message);
                form1.LogFile("DELETE product button : " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }

        }



        //FILL data /////////////////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// FILL Datagridview function
        /// </summary>
        /// <param name="id"></param>
        /// <param name="user"></param>
        /// <param name="pass"></param>
        /// 
        //private void filldgv(string ID, string name, string color, string size, float weight, int count, float price, string Description, string Date_Time)
        //{
        //    MessageBox.Show(ID + name);
        //    dataGridView1.Rows.Add(ID, name, color, size, weight, count, price, Description, Date_Time);
        //}

        private void filldgv(string ID, string name,string color, string size, double weight, int count, double price, string Description, string Date_Time , string image1)
        {
            
            dataGridView1.Rows.Add(ID, name, color, size, weight, count , price , Description, Date_Time,  image1 );
        }



        /// <summary>
        /// retrival(shelfat ntonem) of data function
        /// </summary>
        private void retrieve()
        {
            dataGridView1.Rows.Clear();
            string sql = "SELECT * FROM products ORDER BY [Date_Time] DESC ";
            cmd = new OleDbCommand(sql, con);
            //con.Close(); //close before open again 
            try
            {
                con.Open();
                adapter = new OleDbDataAdapter(cmd);
                adapter.Fill(dt);
                //loop thru dt
                
                foreach (DataRow row in dt.Rows)
                {
   filldgv(row[0].ToString(), row[1].ToString(), row[2].ToString(), row[3].ToString(), Convert.ToDouble(row[4].ToString()), Convert.ToInt32(row[5].ToString()), Convert.ToDouble(row[6].ToString()), row[7].ToString(), row[8].ToString(), row[9].ToString());

                }

                //clear dt 
                dt.Rows.Clear();
                adapter.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show("retrieval: " + ex.Message);
                form1.LogFile("retrival(shelfat ntonem) of data function : " + ex.Message, ex.StackTrace, this.FindForm().Name);

            }
            finally { con.Close(); }
        }

        private void getimages()
        {
            dataGridView1.Rows.Clear();
            string sql = "SELECT ID,Date_Time,image1 FROM products ORDER BY [Date_Time] DESC ";
            cmd = new OleDbCommand(sql, con);
            //con.Close(); //close before open again 
            try
            {
                con.Open();
                adapter = new OleDbDataAdapter(cmd);
                adapter.Fill(dt);
                //loop thru dt

                foreach (DataRow row in dt.Rows)
                {
                    //;filldgv(row[0].ToString(), row[8].ToString(), row[9].ToString());

                }

                //clear dt 
                dt.Rows.Clear();
                adapter.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show("retrieval: " + ex.Message);
                form1.LogFile("retrival(shelfat ntonem) of data function : " + ex.Message, ex.StackTrace, this.FindForm().Name);

            }
            finally { con.Close(); }
        }
  


        //clear and exit//////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// clear txt function
        /// </summary>
        private void clearTxts()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            pictureBox1.Image = null;
            label9.Text = "";
        }


        /// <summary>
        /// clear textboxes and datagridview
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
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



        private void product_Load(object sender, EventArgs e)
        {

        }



        ////textbox RegularExpressions /////////////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// textbox1 (name)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox1.Text, "[^a-zA-Z]"))
            {
                ///place cursor at the end,save position textbox at the end after remove length-1
                bool focused = textBox1.Focused;
                int start = textBox1.SelectionStart;
                int len = textBox1.SelectionLength;

                //remove the wrong digit
                MessageBox.Show("Enter a valid Name A-Z ");             
                textBox1.Text = textBox1.Text.Remove(textBox1.Text.Length - 1);

                ///restore to the last position cursor
                textBox1.SelectionStart = start;
                textBox1.SelectionLength = len;
                textBox1.Select();

                //log file
                form1.LogFile(sender + "-product Textbox1(name):" + textBox1.Text, ((Control)sender).Name, this.FindForm().Name);
            }
        }


        /// <summary>
        /// textbox2 (color)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox2.Text, "[^a-zA-Z]"))
            {
                ///place cursor at the end,save position textbox at the end after remove length-1
                bool focused = textBox2.Focused;
                int start = textBox2.SelectionStart;
                int len = textBox2.SelectionLength;

                //remove the wrong digit
                MessageBox.Show("Enter a valid Name A-Z ");
                textBox2.Text = textBox2.Text.Remove(textBox2.Text.Length - 1);

                ///restore to the last position cursor
                textBox2.SelectionStart = start;
                textBox2.SelectionLength = len;
                textBox2.Select();

                //log file
                form1.LogFile(sender + "-product Textbox2(color):" + textBox2.Text, ((Control)sender).Name, this.FindForm().Name);
            }
        }


        /// <summary>
        /// textbox3 (size)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox3.Text, "[^0-9a-zA-Z]"))
            {
                ///place cursor at the end,save position textbox at the end after remove length-1
                bool focused = textBox3.Focused;
                int start = textBox3.SelectionStart;
                int len = textBox3.SelectionLength;

                //remove the wrong digit
                MessageBox.Show("Enter a valid Name/Number 0-9 , A-Z ");
                textBox3.Text = textBox3.Text.Remove(textBox3.Text.Length - 1);

                ///restore to the last position cursor
                textBox3.SelectionStart = start;
                textBox3.SelectionLength = len;
                textBox3.Select();

                //log file
                form1.LogFile(sender + "-product Textbox3(size):" + textBox3.Text, ((Control)sender).Name, this.FindForm().Name);
            }
        }


        /// <summary>
        /// textbox4 (Weight)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox4.Text, @"[^0-9\.]"))

            {
                ///place cursor at the end,save position textbox at the end after remove length-1
                bool focused = textBox4.Focused;
                int start = textBox4.SelectionStart;
                int len = textBox4.SelectionLength;

                //remove the wrong digit
                MessageBox.Show("Enter a valid Number");
                textBox4.Text = textBox4.Text.Remove(textBox4.Text.Length - 1);

                ///restore to the last position cursor
                textBox4.SelectionStart = start;
                textBox4.SelectionLength = len;
                textBox4.Select();

                //log file
                form1.LogFile(sender + "-product Textbox4(Weight):" + textBox4.Text, ((Control)sender).Name, this.FindForm().Name);
            }
        }


        /// <summary>
        /// textbox5 (count)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox5.Text, "[^0-9]"))
            {
                ///place cursor at the end,save position textbox at the end after remove length-1
                bool focused = textBox5.Focused;
                int start = textBox5.SelectionStart;
                int len = textBox5.SelectionLength;

                //remove the wrong digit
                MessageBox.Show("Enter a valid Number");
                textBox5.Text = textBox5.Text.Remove(textBox5.Text.Length - 1);
                
                ///restore to the last position cursor
                textBox5.SelectionStart = start;
                textBox5.SelectionLength = len;
                textBox5.Select();

                //logfile
                form1.LogFile(sender + "-product Textbox5(count):" + textBox5.Text, ((Control)sender).Name, this.FindForm().Name);
            }
        }


        /// <summary>
        /// textbox6 (price)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox6_TextChanged(object sender, EventArgs e)
        {         
          if (System.Text.RegularExpressions.Regex.IsMatch(textBox6.Text, "[^0-9]"))
             {
                ///place cursor at the end,save position textbox at the end after remove length-1
                bool focused = textBox6.Focused;
                int start = textBox6.SelectionStart;
                int len = textBox6.SelectionLength;

                //remove the wrong digit
                MessageBox.Show("Enter a valid Number");
                textBox6.Text = textBox6.Text.Remove(textBox6.Text.Length - 1);

                ///restore to the last position cursor
                textBox6.SelectionStart = start;
                textBox6.SelectionLength = len;
                textBox6.Select();

                //logfile
                form1.LogFile(sender + "-product Textbox6(price):" + textBox6.Text, ((Control)sender).Name, this.FindForm().Name);
            }        
        }


        /// <summary>
        /// export datagridview to pdf file ... pdf button 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        int counter;
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
                form1.LogFile("PDF product button-Exception-  " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }

            //Exporting to PDF
            //string folderPath = "C:\\PDFs\\";

            string folderPath = Directory.GetCurrentDirectory() + @"\PDFs\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            using (FileStream stream = new FileStream(folderPath +"DataGridViewExport"+counter+".pdf", FileMode.Create))
            {
                Document pdfDoc = new Document(PageSize.A2, 10f, 10f, 10f, 0f);
                PdfWriter.GetInstance(pdfDoc, stream);
                pdfDoc.Open();
                pdfDoc.Add(pdfTable);
                pdfDoc.Close();
                stream.Close();
                counter++; //file0,file1,file2 - add num the name of the file 
                MessageBox.Show("exported ...");
            }
            
        }



        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }





        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }




    }
}
