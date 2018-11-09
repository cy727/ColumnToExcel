using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ColumnToExcel
{
    public partial class PrintOptions : Form
    {
        public PrintOptions()
        {
            InitializeComponent();
        }
        public PrintOptions(List<string> availableFields,string strTitle,bool bPrv)
        {
            InitializeComponent();

            foreach (string field in availableFields)
                     chklst.Items.Add(field, true);

                 txtTitle.Text = strTitle;

                 if (bPrv)
                     checkBoxPrv.Checked = true;
                 else
                     checkBoxPrv.Checked = false;
        }

        private void PrintOtions_Load(object sender, EventArgs e)
        {
            // Initialize some controls
            rdoAllRows.Checked = true;
            chkFitToPageWidth.Checked = true;
            chkWarn.Checked = false;
            this.TopMost = true;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        public List<string> GetSelectedColumns()
        {
            List<string> lst = new List<string>();
            foreach (object item in chklst.CheckedItems)
                    lst.Add(item.ToString());
            return lst;
        }

        public string PrintTitle
        {
            get { return txtTitle.Text; }
        }

        public bool PrintAllRows
        {
            get { return rdoAllRows.Checked; }
        }

        public bool FitToPageWidth
        {
            get { return chkFitToPageWidth.Checked; }
        }

        public bool PrintPreView
        {
            get { return checkBoxPrv.Checked; }
        }

        public bool PrintToFile
        {
            get { return  checkBoxExcel.Checked; }
        }

        public bool PrintWarn
        {
            get { return chkWarn.Checked; }
        }
        public bool PrintXH
        {
            get { return checkBoxXH.Checked; }
        }

    }
}