using ADOX;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DemoDataGridView1
{
    public partial class Form3 : Form
    {
        private List<String> listOfCourseIDs;
        private int page = 0;
        private int NumberOfPage = 0;
        private DataTable dataTable;
        private int NumberOfSlot = 5;

        public Form3()
        {
            InitializeComponent();
        }

        private void generateRows()
        {
            while (dataGridView1.Rows.Count > 0)
            {
                dataGridView1.Rows.RemoveAt(0);
            }
            for (int i = 0; i < NumberOfSlot; i++)
            {
                dataGridView1.Rows.Add("Slot " + (i + 1));
            }
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            DemoDataGridView2.ChangeSlot form = new DemoDataGridView2.ChangeSlot();
            DialogResult dialogResult = form.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                NumberOfSlot = form.slot;
                // Console.WriteLine(NumberOfSlot);
            }
            listOfCourseIDs = GetSheetNames(Directory.GetCurrentDirectory() + "\\DetailExam.xlsx");
            dataGridView1.Columns.Add("title", "Slot/Day");
            dataGridView1.Columns.Add("dayOne", "Day 1");
            dataGridView1.Columns.Add("dayTwo", "Day 2");
            dataGridView1.Columns.Add("dayThree", "Day 3");
            dataGridView1.Columns.Add("dayFour", "Day 4");
            dataGridView1.Columns.Add("dayFive", "Day 5");
            dataGridView1.Columns.Add("daySix", "Day 6");
            dataGridView1.Columns.Add("daySeven", "Day 7");

            generateRows();

            dataTable = new DataTable();
            dataTable.Columns.Add("Course ID");
            dataTable.Columns.Add("Course Name");
            dataTable.Columns.Add("Day");
            dataTable.Columns.Add("Slot");
            dataTable.Columns.Add("Room ID");
            dataTable.Columns.Add("Short Text");
            dataTable.Columns.Add("Long Text");

            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

            this.Enabled = false;
            NumberOfPage = 0;
            page = 1;
            pageLabel.Text = "/";
            bool checker = false;
            //List<String> Test = GetSheetNames(Directory.GetCurrentDirectory() + "\\DanhSachMonHoc.xlsx");
            DataTable listOfCourse = GetAllDataFromSheet(Directory.GetCurrentDirectory() + "\\DanhSachMonHoc.xlsx", "DanhSachMonHoc");
            foreach (String courseID in listOfCourseIDs)
            {
                try
                {
                    DataTable result = GetDataFromSheet(Directory.GetCurrentDirectory() + "\\DetailExam.xlsx", courseID);
                    // if (null == dataTable)
                    //{
                    // dataTable = result;
                    //}
                    //else
                    //{
                    String courseName = "";
                    foreach (DataRow row in listOfCourse.Rows)
                    {
                        String id = row[0].ToString();
                        if (id.Equals(courseID))
                        {
                            courseName = row[1].ToString();
                            break;
                        }
                    }
                    String shortText = courseID;// +": ";
                    String longText = courseName + "(MãMH:" + courseID + ")" + Environment.NewLine;
                    int day = -1;
                    int slot = -1;
                    String roomID = "";
                    List<String> roomIDs = new List<string>();
                    foreach (DataRow row in result.Rows)
                    {
                        // column 0 mssv
                        // column 1 lop
                        // column 2 ho lot
                        // column 3 ten
                        // column 4 ngay thi
                        day = Int16.Parse(row[5].ToString());
                        // column 5 ca thi
                        slot = Int16.Parse(row[6].ToString());
                        // column 6 phong thi
                        roomID = row[7].ToString();
                        if (!roomIDs.Contains(roomID))
                        {
                            roomIDs.Add(roomID);
                        }
                        // }
                    }
                    if (day > -1 && slot > -1)
                    {
                        //dataGridView1.Rows[slot].Cells[day].Value = text;
                        String roomString = "";
                        foreach (String room in roomIDs)
                        {
                            roomString += room + " - ";
                        }
                        roomString = roomString.Substring(0, roomString.Length - 2);
                        //shortText += " " + roomString + Environment.NewLine;
                        longText += "Phòng " + roomString + Environment.NewLine;

                        DataRow newRow = dataTable.NewRow();
                        newRow.SetField<String>("Course ID", courseID);
                        newRow.SetField<String>("Course Name", courseName);
                        newRow.SetField<String>("Day", day.ToString());
                        newRow.SetField<String>("Slot", slot.ToString());
                        newRow.SetField<String>("Room ID", roomString);
                        newRow.SetField<String>("Short Text", shortText);
                        newRow.SetField<String>("Long Text", longText);
                        dataTable.Rows.Add(newRow);
                        if (day > 7)
                        {
                            page = 1;
                            int value = day / 7 + 1;
                            if (value > NumberOfPage)
                                NumberOfPage = value;
                        }
                        checker = true;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(courseID + ": " + ex.Message, courseID + " Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (checker)
            {
                if (NumberOfPage > 1)
                {
                    pageLabel.Text = "1/2";
                    previousButton.Enabled = false;
                    nextButton.Enabled = true;
                }
                else
                {
                    pageLabel.Text = "1/1";
                    previousButton.Enabled = false;
                    nextButton.Enabled = false;
                }

                foreach (DataRow row in dataTable.Rows)
                {
                    int day = Int16.Parse(row[2].ToString());
                    int slot = Int16.Parse(row[3].ToString());
                    if (page < 2 && day <= 7)
                    {
                        dataGridView1.Rows[slot - 1].Cells[day].Value += row[5].ToString();
                        dataGridView1.Rows[slot - 1].Cells[day].Value += "\r\n";
                        dataGridView1.Rows[slot - 1].Cells[day].ToolTipText += row[6].ToString();
                        dataGridView1.Rows[slot - 1].Cells[day].ToolTipText += "\r\n";
                    }
                }
                //Form2 form2 = new Form2(ID);
                //form2.Show();
            }
            else
            {
                MessageBox.Show("Student's ID doesn't exist !");
                nextButton.Enabled = false;
                previousButton.Enabled = false;
            }
            this.Enabled = true;
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

        public static DataTable GetDataFromSheet(string strFileName, string sheetname)
        {
            System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " + strFileName + "; Extended Properties = \"Excel 8.0;HDR=Yes;IMEX=1\";");
            string strQuery = "SELECT * FROM [\'" + sheetname + "$\']";
            OleDbCommand command = new OleDbCommand(strQuery, conn);
            conn.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter(command);
            System.Data.DataSet ds = new System.Data.DataSet();
            adapter.Fill(ds);
            conn.Close();
            return ds.Tables[0];
        }

        public static DataTable GetAllDataFromSheet(string strFileName, string sheetname)
        {
            System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " + strFileName + "; Extended Properties = \"Excel 8.0;HDR=Yes;IMEX=1\";");
            string strQuery = "SELECT * FROM [" + sheetname + "$]";
            OleDbCommand command = new OleDbCommand(strQuery, conn);
            conn.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter(command);
            System.Data.DataSet ds = new System.Data.DataSet();
            adapter.Fill(ds);
            conn.Close();
            return ds.Tables[0];
        }

        private void nextButton_Click(object sender, EventArgs e)
        {
            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Rows.Count; j++)
                {
                    dataGridView1.Rows[j].Cells[i].Value = DBNull.Value;
                    dataGridView1.Rows[j].Cells[i].ToolTipText = "";
                }
            }
            page++;
            if (page == NumberOfPage)
            {
                nextButton.Enabled = false;
            }
            previousButton.Enabled = true;
            dataGridView1.Columns[1].HeaderText = "Day " + ((page - 1) * 7 + 1);
            dataGridView1.Columns[2].HeaderText = "Day " + ((page - 1) * 7 + 2);
            dataGridView1.Columns[3].HeaderText = "Day " + ((page - 1) * 7 + 3);
            dataGridView1.Columns[4].HeaderText = "Day " + ((page - 1) * 7 + 4);
            dataGridView1.Columns[5].HeaderText = "Day " + ((page - 1) * 7 + 5);
            dataGridView1.Columns[6].HeaderText = "Day " + ((page - 1) * 7 + 6);
            dataGridView1.Columns[7].HeaderText = "Day " + ((page - 1) * 7 + 7);

            foreach (DataRow row in dataTable.Rows)
            {
                int day = Int16.Parse(row[2].ToString());
                int slot = Int16.Parse(row[3].ToString());
                if (page > 1 && day > 7)
                {
                    dataGridView1.Rows[slot - 1].Cells[day - (page - 1) * 7].Value += row[5].ToString();
                    dataGridView1.Rows[slot - 1].Cells[day - (page - 1) * 7].Value += "\r\n";
                    dataGridView1.Rows[slot - 1].Cells[day - (page - 1) * 7].ToolTipText += row[6].ToString();
                    dataGridView1.Rows[slot - 1].Cells[day - (page - 1) * 7].ToolTipText += "\r\n";
                }
            }
        }

        private void previousButton_Click(object sender, EventArgs e)
        {
            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Rows.Count; j++)
                {
                    dataGridView1.Rows[j].Cells[i].Value = DBNull.Value;
                    dataGridView1.Rows[j].Cells[i].ToolTipText = "";
                }
            }
            page--;
            if (page == 1)
            {
                previousButton.Enabled = false;
            }
            nextButton.Enabled = true;
            dataGridView1.Columns[1].HeaderText = "Day " + ((page - 1) * 7 + 1);
            dataGridView1.Columns[2].HeaderText = "Day " + ((page - 1) * 7 + 2);
            dataGridView1.Columns[3].HeaderText = "Day " + ((page - 1) * 7 + 3);
            dataGridView1.Columns[4].HeaderText = "Day " + ((page - 1) * 7 + 4);
            dataGridView1.Columns[5].HeaderText = "Day " + ((page - 1) * 7 + 5);
            dataGridView1.Columns[6].HeaderText = "Day " + ((page - 1) * 7 + 6);
            dataGridView1.Columns[7].HeaderText = "Day " + ((page - 1) * 7 + 7);

            foreach (DataRow row in dataTable.Rows)
            {
                int day = Int16.Parse(row[2].ToString());
                int slot = Int16.Parse(row[3].ToString());
                if (page < 2 && day <= 7)
                {
                    dataGridView1.Rows[slot - 1].Cells[day - (page - 1) * 7].Value += row[5].ToString();
                    dataGridView1.Rows[slot - 1].Cells[day].Value += "\r\n";
                    dataGridView1.Rows[slot - 1].Cells[day - (page - 1) * 7].ToolTipText += row[6].ToString();
                    dataGridView1.Rows[slot - 1].Cells[day].ToolTipText += "\r\n";
                }
            }
        }

        private void dataGridView1_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= 1 & e.RowIndex >= 1)
            {
               
            }
        }
    }
}
