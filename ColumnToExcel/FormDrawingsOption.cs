using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ColumnToExcel
{
    public partial class FormDrawingsOption : Form
    {
        public bool bCancel = false;
        public Font fontCatia = new System.Drawing.Font("Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
        public Font fontCatia1 = new System.Drawing.Font("Microsoft Sans Serif", 10, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
        public Font fontCatia2 = new System.Drawing.Font("Microsoft Sans Serif", 6, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);

        public FormDrawingsOption()
        {
            InitializeComponent();
            comboBoxZZ.SelectedIndex = 0;
            comboBoxFH.SelectedIndex = 4;
        }

        private void FormDrawingsOption_Load(object sender, EventArgs e)
        {
            this.TopMost = true;
            //comboBoxZZ.SelectedIndex = 0;

            showZT(0);
            showZT(1);
            showZT(2);
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
                showZT(0);
            }
        }

        private void buttonZT1_Click(object sender, EventArgs e)
        {
            if (fontDialogCatia1.ShowDialog() == DialogResult.OK)
            {
                fontCatia1 = fontDialogCatia1.Font;
                showZT(1);
            }
        }

        private void showZT(int iZT)
        {
            Font fT;
            string sT;


            switch (iZT)
            {
                case 0:
                    fT = fontCatia;
                    break;
                case 1:
                    fT = fontCatia1;
                    break;
                case 2:
                    fT = fontCatia2;
                    break;
                default:
                    return;
            }

            sT = fT.Name+" Size:"+fT.Size.ToString();

            if (fT.Bold)
                sT += " Bold";
            if (fT.Underline)
                sT += " Underline";
            if (fT.Italic)
                sT += " Italic";

            switch (iZT)
            {
                case 0:
                    textBoxZT.Text=sT;
                    break;
                case 1:
                    textBoxZT1.Text = sT;
                    break;
                case 2:
                    textBoxZT2.Text = sT;
                    break;
                default:
                    return;
            }


        }

        private void buttonZT2_Click(object sender, EventArgs e)
        {
            if (fontDialogCatia2.ShowDialog() == DialogResult.OK)
            {
                fontCatia2 = fontDialogCatia2.Font;
                showZT(2);
            }
        }
    }
}
