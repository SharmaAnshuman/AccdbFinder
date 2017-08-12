using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Threading;
using System.IO;
using ADOX;
namespace Finder
{
    public partial class Form1 : Form
    {

        DataSet ds;
        static String[] tablenm = new String[] { "MSysAccessStorage", "MSysACEs", "MSysComplexColumns", "MSysNameMap", "MSysNavPaneGroupCategories", "MSysNavPaneGroups", "MSysNavPaneGroupToObjects", "MSysNavPaneObjectIDs", "MSysObjects", "MSysQueries", "MSysResources", "MSysRelationships", "MSysAccessXML" };
        string sql = null;
        static String connetionString = null;
        static String connetionString1 = null;
        static OleDbConnection oledbCnn, oledbCnn1;
        static OleDbDataAdapter oledbAdp;
        String tblnm="";
        static TextBox[] textBox = new TextBox[99];
        static Label[] lable = new Label[99];
        static int totalTxtGen = 0;
        static int datalimit = 3000;
        static Form2 x;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }



        private void openDatabaseFunction1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            flowLayoutPanel1.Controls.Clear();
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    String filename = openFileDialog.FileName;
                    Thread thread = new Thread((object selectedFilename) =>
                    {
                        try
                        {
                            connetionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + selectedFilename + ";";
                            oledbCnn = new OleDbConnection(connetionString);
                            oledbCnn.Open();
                            DataTable dt123 = oledbCnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            Invoke(new Action(() =>
                            {
                                comboBox2.Items.Clear();
                                //comboBox2.SelectedIndex = -1;
                            }));
                            foreach (DataRow row in dt123.Rows)
                            {
                                if (!(tablenm.Contains(row["TABLE_NAME"].ToString())))
                                {
                                    Invoke(new Action(() => {
                                        comboBox2.Items.Add(row["TABLE_NAME"].ToString());
                                    }));
                                }
                            }
                            Invoke(new Action(() => {
                                //comboBox2.SelectedItem = comboBox2.Items[0];
                            }));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);
                        }
                        Invoke(new Action(() =>
                        {
                            x.Close();
                        }));
                    });
                    thread.Start(filename);
                    x = new Form2();
                    x.ShowDialog();
                }
            }
            catch (Exception ex1) { }
        }

      /*  private void openDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Thread t1 = new Thread(wait);
            
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    t1.Start();
                    Action action = () =>
                    {
                        connetionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + openFileDialog.FileName.ToString() + ";";
                        oledbCnn = new OleDbConnection(connetionString);
                        oledbCnn.Open();
                        DataTable dt123 = oledbCnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        comboBox2.Items.Clear();
                        foreach (DataRow row in dt123.Rows)
                        {
                            if (!(tablenm.Contains(row["TABLE_NAME"].ToString())))
                            {
                                comboBox2.Items.Add(row["TABLE_NAME"].ToString());
                            }
                        }
                        comboBox2.SelectedItem = comboBox2.Items[0];
                        x.Close();
                    };
                    Invoke(action);
                }
                
            }
            catch (Exception x)
            {
                MessageBox.Show(x.ToString());
            }
            comboBox2.Focus();
        }
        */
        
        public static void wait()
        {
            x = new Form2();
            x.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
            try
            {
                if (connetionString != null)
                {
                    try
                    {
                        if (checkBox1.Checked != true)
                        {
                            sql = getQry();
                        }
                        else
                        {
                            sql = getQry();
                            int totalLen = sql.Length;
                            int whereLen = sql.LastIndexOf("where");
                            if ((totalLen - whereLen) != 6)
                            {
                                sql = sql.Replace("where", "where NOT(");
                            }

                            if ((totalLen - whereLen) == 6)
                            {
                                sql = sql.Replace("where", " ");
                            }
                            else
                            {
                                sql += ")";
                            }

                        }
                        Thread thread = new Thread(() =>
                            {
                                try
                                {

                                    oledbCnn = new OleDbConnection(connetionString);
                                    oledbCnn.Open();
                                    sql = sql.Replace("*", "TOP " + datalimit + " *");
                                    if (checkBox1.Checked == true)
                                        sql = sql.Replace(")", "");
                                    oledbAdp = new OleDbDataAdapter(sql, oledbCnn);
                                    ds = new DataSet();
                                    oledbAdp.Fill(ds, tblnm);
                                    DataTable dt1 = new DataTable();
                                    dt1 = ds.Tables[0];
                                    Invoke(new Action(() =>
                                    {
                                        dataGridView1.Refresh();
                                        dataGridView1.DataSource = ds;
                                        dataGridView1.DataMember = tblnm;
                                    }));
                                    //textBox[totalTxtGen - 1].Text = sql;
                                    oledbCnn.Close();
                                }
                                catch (Exception x)
                                {

                                    MessageBox.Show(x + "");
                                }
                                finally
                                {
                                    Invoke(new Action(() =>
                                    {
                                        x.Close();
                                    }));
                                }
                            });
                        thread.Start();
                        x = new Form2();
                        x.ShowDialog();

                    }
                    catch (Exception x)
                    {
                        MessageBox.Show("Please Select Table.!");

                    }
                }
                else
                {
                    MessageBox.Show("Goto File Menu Open Database");
                }
            }
            catch(Exception x)
            {
                MessageBox.Show(x.Message.ToString());
            }
         }

        private void button2_Click(object sender, EventArgs e)
        {

            if (ds != null && tblnm != null && tblnm != "" && tblnm != " ")
            {
                SaveFileDialog a = new SaveFileDialog();
                a.Filter = "Access Database File (*.accdb)|*.accdb|All files (*.*)|*.*";
                if (a.ShowDialog() == DialogResult.OK)
                {
                    String saveFilenm = a.FileName.ToString();
                    ADOX.Catalog cat = new ADOX.Catalog();
                    try
                    {
                        Thread thread = new Thread(() =>
                        {
                            try
                            {

                                cat.Create(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + saveFilenm);
                                connetionString1 = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + saveFilenm + ";Jet OLEDB:Database Password=Pass123;";
                                oledbCnn1 = new OleDbConnection(connetionString1);
                                oledbCnn1.Open();
                                String strQury = "create table `" + tblnm + "`  (";
                                String col = "", para = "";
                                int i1;
                                for (i1 = 0; i1 < totalTxtGen + 1; i1++)
                                {
                                    if (i1 == totalTxtGen)
                                    {
                                        col += "`" + textBox[i1].Name + "` TEXT ";
                                    }
                                    else
                                    {
                                        col += "`" + textBox[i1].Name + "` TEXT ,";
                                    }
                                    para += "@" + textBox[i1].Name + "~";
                                }
                                //para += "@" + textBox[i1].Name + "~";
                                strQury += col + " )";


                                strQury = strQury.Replace('.', ' ');
                                OleDbCommand cmd = new OleDbCommand(strQury, oledbCnn1);
                                cmd.ExecuteNonQuery();
                                String[] p = para.Split('~');
                                String tmp1 = "insert into `" + tblnm + "` (";

                                for (int i = 0; i < (p.Length - 1); i++)
                                {
                                    if (i == p.Length - 2)
                                        tmp1 += "`" + p[i].Replace('@', ' ').Trim() + "` ";
                                    else
                                        tmp1 += "`" + p[i].Replace('@', ' ').Trim() + "` ,";
                                }


                                tmp1 += ") VALUES (";
                                for (int i = 0; i < (p.Length - 1); i++)
                                {
                                    if (i == p.Length - 2)
                                        tmp1 += "`" + p[i] + "`";
                                    else
                                        tmp1 += "`" + p[i] + "` ,";
                                }
                                tmp1 += ")";

                                if (oledbCnn.State == ConnectionState.Closed)
                                {
                                    oledbCnn.Open();
                                }
                                cmd = new OleDbCommand(getQry(), oledbCnn);
                                OleDbDataReader dr = cmd.ExecuteReader();
                                OleDbCommand cmd1 = new OleDbCommand(tmp1, oledbCnn1);
                                while (dr.Read())
                                {
                                    cmd1.Parameters.Clear();
                                    for (int co = 0; co < p.Length - 1; co++)
                                    {
                                        cmd1.Parameters.AddWithValue(p[co] + "", dr.GetValue(co).ToString());
                                    }
                                    cmd1.ExecuteNonQuery();
                                }

                                cmd1.Dispose();
                                oledbCnn1.Close();
                                MessageBox.Show("File Saved..");

                            }
                            catch (System.Runtime.InteropServices.COMException x)
                            {
                                MessageBox.Show(x + "");
                            }
                            finally
                            {
                                Invoke(new Action(() =>
                                {
                                    x.Close();
                                }));
                            }

                        });
                        thread.Start();
                        x = new Form2();
                        x.ShowDialog();
                    }
                    catch (Exception x)
                    {
                        MessageBox.Show("" + x);
                    }

                }
                else
                {
                    MessageBox.Show("Please Open Database First.");
                }

            }
            else
            {
                MessageBox.Show("Please Open Database First.");
            }
        }

        public TextBox addtextbox(String x)
        {
            TextBox a = new TextBox();
            a.Name = x;
            a.Width = 300;
            return a;
        }
        public Label addlabel(String x)
        {
            Label a = new Label();
            a.Name = x;
            a.Width = 120;
            a.Text = x;
            return a;
        }
        private static string MySQLEscape(string str)
        {
           
            return Regex.Replace(str, @"[\x00'""\b\n\r\t\cZ\\%_]",
                delegate(Match match)
                {
                    string v = match.Value;
                    switch (v)
                    {
                        case "\'":
                            return" ";
                        case "\x00":            // ASCII NUL (0x00) character
                            return "\\0";
                        case "\b":              // BACKSPACE character
                            return "\\b";
                        case "\n":              // NEWLINE (linefeed) character
                            return "\\n";
                        case "\r":              // CARRIAGE RETURN character
                            return "\\r";
                        case "\t":              // TAB
                            return "\\t";
                        case "\u001A":          // Ctrl-Z
                            return "\\Z";
                        default:
                            return "\\" + v;
                    }
                });
        }

        private void saveDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button2_Click(sender, e);
        }
        
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 ab = new AboutBox1();
            ab.ShowDialog();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void comboBox2_selectedIndexChangedNew(object sender, EventArgs e)
        {
            try
            {
                tblnm = comboBox2.Text.ToString();
                Thread thread = new Thread((object tableName) => {
                    try
                    {
                        // 
                        oledbAdp = new OleDbDataAdapter("select top 1 * from `" + tableName + "`", oledbCnn);
                        ds = new DataSet();
                        oledbAdp.Fill(ds, tableName.ToString());
                        DataTable dt1 = new DataTable();
                        dt1 = ds.Tables[0];

                        int j = 0;
                        Invoke(new Action(() => {
                            flowLayoutPanel1.WrapContents = true;
                            flowLayoutPanel1.AutoScroll = true;
                            flowLayoutPanel1.Controls.Clear();
                        }));

                        foreach (DataColumn column in dt1.Columns)
                        {
                            Label c1 = addlabel(column.ColumnName.ToString());
                            lable[j] = c1;

                            TextBox c = addtextbox(column.ColumnName.ToString());
                            textBox[j] = c;

                            j++;

                            Invoke(new Action(() => {
                                flowLayoutPanel1.Controls.Add(c1);
                                flowLayoutPanel1.Controls.Add(c);
                            }));
                        }
                        totalTxtGen = j - 1;

                        Invoke(new Action(() =>
                        {
                            x.Close();
                        }));
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(""+ex);
                        Invoke(new Action(() =>
                        {
                            x.Close();
                        }));
                    }
                    Invoke(new Action(() => {
                        x.Close();
                    }));
                });
                thread.Start(tblnm);
                x = new Form2();
                x.ShowDialog();
            }
            catch (Exception x)
            {
                MessageBox.Show(x.GetBaseException().ToString());
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                tblnm = comboBox2.Text.ToString();
                oledbAdp = new OleDbDataAdapter("select * from `" + tblnm + "`", oledbCnn);
                ds = new DataSet();
                oledbAdp.Fill(ds, tblnm);
                DataTable dt1 = new DataTable();
                dt1 = ds.Tables[0];

                int j = 0;
                flowLayoutPanel1.WrapContents = true;
                flowLayoutPanel1.AutoScroll = true;
                flowLayoutPanel1.Controls.Clear();
                
                foreach (DataColumn column in dt1.Columns)
                {
                    Label c1 = addlabel(column.ColumnName.ToString());
                    flowLayoutPanel1.Controls.Add(c1);
                    lable[j] = c1;

                    TextBox c = addtextbox(column.ColumnName.ToString());
                    flowLayoutPanel1.Controls.Add(c);
                    textBox[j] = c;
             
                    j++;
                    
                }
                totalTxtGen = j - 1;
            }
            catch (Exception x)
            {
                MessageBox.Show(x.GetBaseException().ToString());
            }
        }
        
        public String getQry()
        {
            String tmp = "", q = "select * from `" + tblnm+"`";
            if (connetionString != null)
            {
                for (int i = 0; i <= totalTxtGen; i++)
                {
                    if (q.EndsWith("`"+tblnm+"`"))
                    {
                        q += " where ";
                    }
                    if (textBox[i].Text != "" && textBox[i].Text != null)
                    {
                        if (textBox[i].Text.Contains("@"))
                        {
                            String[] str = textBox[i].Text.Split('@');
                            tmp += "(";
                            for (int i1 = 0; i1 < str.Length; i1++)
                            {
                                
                                if (i1 == (str.Length - 1))
                                {
                                    tmp += " `"+textBox[i].Name.ToString()+"` like '%" + str[i1] + "%' ";
                                }
                                else
                                {
                                    tmp += "`" + textBox[i].Name.ToString() + "` like '%" + str[i1] + "%' or ";
                                }
                            }
                            tmp += ")";
                        }
                        else
                        {
                            tmp += "(`" + textBox[i].Name.ToString() + "` like '%" + textBox[i].Text.ToString() + "%')";
                        }
                    }
                }
            }
            
            q += tmp;
            q = q.Replace(")(", ") and (");

            int totalLen = q.Length;
            int whereLen = q.LastIndexOf("where");
            if ((totalLen - whereLen) == 6)
            {
                q = q.Replace("where", " ");
            }

            return(q);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            /*

            try
            {

                if (ds != null)
                {
                    SaveFileDialog a = new SaveFileDialog();
                    a.Filter = "CSV file (*.csv)|*.csv| All Files (*.*)|*.*";
                    if (a.ShowDialog() == DialogResult.OK)
                    {
                        Thread thread = new Thread(() =>
                        {
                            String saveFilenm = a.FileName.ToString();
                            try
                            {
                                FileStream fs1 = new FileStream(a.FileName.ToString(), FileMode.OpenOrCreate, FileAccess.Write);
                                StreamWriter writer = new StreamWriter(fs1);
                                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                                {
                                    writer.Write(ds.Tables[0].Rows[i].ItemArray[0] + "," + ds.Tables[0].Rows[i].ItemArray[1] + "," + ds.Tables[0].Rows[i].ItemArray[2].ToString() + "\n");
                                }
                                writer.Close();
                                MessageBox.Show("File Saved..\n" + a.FileName.ToString());
                            }
                            catch (Exception x)
                            {
                                MessageBox.Show(x.Message.ToString());
                            }
                            finally
                            {
                                Invoke(new Action(() =>
                               {
                                   x.Close();
                               }));

                            }

                        });


                        thread.Start();
                        x = new Form2();
                        x.Show();

                    }
                    else
                    {
                        MessageBox.Show("Please Open Data First.");
                    }

                }
            }
            catch(Exception x)
            {
                MessageBox.Show("" + x);
             
              }
             */
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if(!(datalimit<3000))
            datalimit += 3000;
            button1_Click(sender,e);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if(datalimit>3000)
            datalimit -= 3000;
            button1_Click(sender, e);
        }
    }
}
