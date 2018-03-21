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
    public partial class Form2 : Form
    {
        private String courseID;
        public Form2(String courseID)
        {
            InitializeComponent();
            this.courseID = courseID;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.Text = "Result for " + courseID;
            DataTable dt = GetDataFromSheet(Directory.GetCurrentDirectory() + "\\DetailExam.xlsx", courseID);
            dataGridView1.DataSource = dt.DefaultView;
            int width = (dataGridView1.Columns.Count) * dataGridView1.Columns[0].Width + dataGridView1.RowHeadersWidth;
            //dataGridView1.Columns[dataGridView1.Columns.Count - 1].Visible = false;
            int height = dataGridView1.Rows.Count * dataGridView1.Rows[0].Height;
            dataGridView1.Size = new Size(width, (height > Screen.PrimaryScreen.Bounds.Height - 50) ? Screen.PrimaryScreen.Bounds.Height - 80 : height);
            this.Size = new Size(dataGridView1.Size.Width + 40, (height > Screen.PrimaryScreen.Bounds.Height) ? Screen.PrimaryScreen.Bounds.Height : height);
            dataGridView1.TopLeftHeaderCell.Value = "STT";
        }

        public static DataTable GetDataFromSheet(string strFileName, string sheetname)
        {
            System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source = " + strFileName + "; Extended Properties = \"Excel 8.0;HDR=Yes;IMEX=1\";");
            conn.Open();
            string strQuery = "SELECT * FROM [" + sheetname + "$" + "]";
            System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter(strQuery, conn);
            System.Data.DataSet ds = new System.Data.DataSet();
            adapter.Fill(ds);
            return ds.Tables[0];
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIdx = (e.RowIndex + 1).ToString();

            var centerFormat = new StringFormat()
            {
                // right alignment might actually make more sense for numbers
                Alignment = StringAlignment.Center,

                LineAlignment = StringAlignment.Center
            };
            //get the size of the string
            Size textSize = TextRenderer.MeasureText(rowIdx, this.Font);
            //if header width lower then string width then resize
            if (grid.RowHeadersWidth < textSize.Width + 40)
            {
                grid.RowHeadersWidth = textSize.Width + 40;
            }
            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
        }
    }
}
