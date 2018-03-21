using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DemoDataGridView2
{
    public partial class ChangeSlot : Form
    {
        public int slot = 5;
        public ChangeSlot()
        {
            InitializeComponent();
        }

        private void okBtn_Click(object sender, EventArgs e)
        {
            slot = (int)numericUpDown1.Value;
            this.DialogResult = DialogResult.OK;
        }
    }
}
