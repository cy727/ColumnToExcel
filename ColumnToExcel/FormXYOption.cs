using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ColumnToExcel
{
    public partial class FormXYOption : Form
    {
        public bool bCancel = false;
        public Font fontCatia;

        public FormXYOption()
        {
            InitializeComponent();
        }

        private void FormXYOption_Load(object sender, EventArgs e)
        {
            this.TopMost = true;

            fontCatia = new System.Drawing.Font("Microsoft Sans Serif",8, System.Drawing.FontStyle.Regular,System.Drawing.GraphicsUnit.Point);
;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            bCancel = false;
            this.Close();
        }

        private void buttonCANCEL_Click(object sender, EventArgs e)
        {
            bCancel = true;
            this.Close();
        }

        private void buttonZT_Click(object sender, EventArgs e)
        {
            if (fontDialogCatia.ShowDialog() == DialogResult.OK)
            {
                fontCatia = fontDialogCatia.Font;
            }
        }
    }
}
