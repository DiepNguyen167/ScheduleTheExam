using ADOX;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DemoDataGridView1
{
    public partial class Form1 : Form
    {
        private List<String> listOfCourseIDs;
        private static String DEFAULT_TEXT = "Input ID in here";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String ID = textBox1.Text;
            if (ID.Equals("") || ID.Equals(DEFAULT_TEXT))
            {
                MessageBox.Show("Please input course's ID");
            }
            else
            {
                bool checker = false;
                foreach (String courseID in listOfCourseIDs)
                {
                    if (courseID.Equals(ID))
                    {
                        checker = true;
                        break;
                    }
                }
                if (checker)
                {
                    this.Hide();
                    Form2 form2 = new Form2(ID);
                    form2.Show();
                }
                else
                {
                    MessageBox.Show("Course's ID doesn't exist !");
                }
            }
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            if (textBox1.Text.Equals(DEFAULT_TEXT))
            {
                textBox1.Text = "";
                textBox1.ForeColor = Color.Black;
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text.Trim().Equals(""))
            {
                textBox1.Text = DEFAULT_TEXT;
                textBox1.ForeColor = Color.Gray;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            listOfCourseIDs = GetSheetNames(Directory.GetCurrentDirectory() + "\\DetailExam.xlsx");
            textBox1.Text = DEFAULT_TEXT;
        }

        public static List<String> GetSheetNames(string strFileName)
        {
            List<String> sheets = new List<string>();
            Catalog oCatlog = new Catalog();
            ADOX.Table oTable = new ADOX.Table();
            ADODB.Connection oConn = new ADODB.Connection();
            oConn.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " + strFileName + "; Extended Properties = \"Excel 8.0;HDR=Yes;IMEX=1\";", "", "", 0);
            oCatlog.ActiveConnection = oConn;
            if (oCatlog.Tables.Count > 0)
            {
                int item = 0;
                foreach (ADOX.Table tab in oCatlog.Tables)
                {
                    if (tab.Type == "TABLE")
                    {
                        sheets.Add(tab.Name.Trim().Substring(1, tab.Name.Length - 3));
                        item++;
                    }
                }
            }
            return sheets;
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
           
        }
    }
}
