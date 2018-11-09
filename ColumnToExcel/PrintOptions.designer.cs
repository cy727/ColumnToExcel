namespace ColumnToExcel
{
    partial class PrintOptions
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PrintOptions));
            this.rdoSelectedRows = new System.Windows.Forms.RadioButton();
            this.rdoAllRows = new System.Windows.Forms.RadioButton();
            this.chkFitToPageWidth = new System.Windows.Forms.CheckBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.txtTitle = new System.Windows.Forms.TextBox();
            this.gboxRowsToPrint = new System.Windows.Forms.GroupBox();
            this.lblColumnsToPrint = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.chklst = new System.Windows.Forms.CheckedListBox();
            this.checkBoxPrv = new System.Windows.Forms.CheckBox();
            this.checkBoxExcel = new System.Windows.Forms.CheckBox();
            this.printDialogDgv = new System.Windows.Forms.PrintDialog();
            this.printPreviewDialog1 = new System.Windows.Forms.PrintPreviewDialog();
            this.chkWarn = new System.Windows.Forms.CheckBox();
            this.checkBoxXH = new System.Windows.Forms.CheckBox();
            this.gboxRowsToPrint.SuspendLayout();
            this.SuspendLayout();
            // 
            // rdoSelectedRows
            // 
            this.rdoSelectedRows.AutoSize = true;
            this.rdoSelectedRows.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoSelectedRows.Location = new System.Drawing.Point(91, 18);
            this.rdoSelectedRows.Name = "rdoSelectedRows";
            this.rdoSelectedRows.Size = new System.Drawing.Size(77, 16);
            this.rdoSelectedRows.TabIndex = 1;
            this.rdoSelectedRows.TabStop = true;
            this.rdoSelectedRows.Text = "选择内容";
            this.rdoSelectedRows.UseVisualStyleBackColor = true;
            // 
            // rdoAllRows
            // 
            this.rdoAllRows.AutoSize = true;
            this.rdoAllRows.Location = new System.Drawing.Point(9, 18);
            this.rdoAllRows.Name = "rdoAllRows";
            this.rdoAllRows.Size = new System.Drawing.Size(51, 16);
            this.rdoAllRows.TabIndex = 0;
            this.rdoAllRows.TabStop = true;
            this.rdoAllRows.Text = "全部";
            this.rdoAllRows.UseVisualStyleBackColor = true;
            // 
            // chkFitToPageWidth
            // 
            this.chkFitToPageWidth.AutoSize = true;
            this.chkFitToPageWidth.CheckAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.chkFitToPageWidth.Checked = true;
            this.chkFitToPageWidth.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkFitToPageWidth.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.chkFitToPageWidth.Location = new System.Drawing.Point(187, 72);
            this.chkFitToPageWidth.Name = "chkFitToPageWidth";
            this.chkFitToPageWidth.Size = new System.Drawing.Size(78, 17);
            this.chkFitToPageWidth.TabIndex = 21;
            this.chkFitToPageWidth.Text = "宽度充满";
            this.chkFitToPageWidth.UseVisualStyleBackColor = true;
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(184, 115);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(59, 13);
            this.lblTitle.TabIndex = 20;
            this.lblTitle.Text = "打印表头";
            // 
            // txtTitle
            // 
            this.txtTitle.AcceptsReturn = true;
            this.txtTitle.Location = new System.Drawing.Point(184, 130);
            this.txtTitle.Multiline = true;
            this.txtTitle.Name = "txtTitle";
            this.txtTitle.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtTitle.Size = new System.Drawing.Size(176, 71);
            this.txtTitle.TabIndex = 19;
            // 
            // gboxRowsToPrint
            // 
            this.gboxRowsToPrint.Controls.Add(this.rdoSelectedRows);
            this.gboxRowsToPrint.Controls.Add(this.rdoAllRows);
            this.gboxRowsToPrint.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gboxRowsToPrint.Location = new System.Drawing.Point(184, 20);
            this.gboxRowsToPrint.Name = "gboxRowsToPrint";
            this.gboxRowsToPrint.Size = new System.Drawing.Size(173, 39);
            this.gboxRowsToPrint.TabIndex = 18;
            this.gboxRowsToPrint.TabStop = false;
            this.gboxRowsToPrint.Text = "打印内容";
            // 
            // lblColumnsToPrint
            // 
            this.lblColumnsToPrint.AutoSize = true;
            this.lblColumnsToPrint.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblColumnsToPrint.Location = new System.Drawing.Point(8, 8);
            this.lblColumnsToPrint.Name = "lblColumnsToPrint";
            this.lblColumnsToPrint.Size = new System.Drawing.Size(46, 13);
            this.lblColumnsToPrint.TabIndex = 17;
            this.lblColumnsToPrint.Text = "打印列";
            // 
            // btnOK
            // 
            this.btnOK.BackColor = System.Drawing.SystemColors.Control;
            this.btnOK.Cursor = System.Windows.Forms.Cursors.Default;
            this.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnOK.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            this.btnOK.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnOK.Image = ((System.Drawing.Image)(resources.GetObject("btnOK.Image")));
            this.btnOK.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnOK.Location = new System.Drawing.Point(239, 228);
            this.btnOK.Name = "btnOK";
            this.btnOK.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnOK.Size = new System.Drawing.Size(56, 23);
            this.btnOK.TabIndex = 14;
            this.btnOK.Text = "确定";
            this.btnOK.UseVisualStyleBackColor = false;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.SystemColors.Control;
            this.btnCancel.Cursor = System.Windows.Forms.Cursors.Default;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnCancel.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            this.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnCancel.Image = ((System.Drawing.Image)(resources.GetObject("btnCancel.Image")));
            this.btnCancel.Location = new System.Drawing.Point(301, 228);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnCancel.Size = new System.Drawing.Size(56, 23);
            this.btnCancel.TabIndex = 15;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // chklst
            // 
            this.chklst.CheckOnClick = true;
            this.chklst.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chklst.FormattingEnabled = true;
            this.chklst.Location = new System.Drawing.Point(8, 26);
            this.chklst.Name = "chklst";
            this.chklst.Size = new System.Drawing.Size(170, 214);
            this.chklst.TabIndex = 13;
            // 
            // checkBoxPrv
            // 
            this.checkBoxPrv.AutoSize = true;
            this.checkBoxPrv.Location = new System.Drawing.Point(262, 72);
            this.checkBoxPrv.Name = "checkBoxPrv";
            this.checkBoxPrv.Size = new System.Drawing.Size(72, 16);
            this.checkBoxPrv.TabIndex = 22;
            this.checkBoxPrv.Text = "打印预览";
            this.checkBoxPrv.UseVisualStyleBackColor = true;
            // 
            // checkBoxExcel
            // 
            this.checkBoxExcel.AutoSize = true;
            this.checkBoxExcel.Location = new System.Drawing.Point(187, 207);
            this.checkBoxExcel.Name = "checkBoxExcel";
            this.checkBoxExcel.Size = new System.Drawing.Size(84, 16);
            this.checkBoxExcel.TabIndex = 23;
            this.checkBoxExcel.Text = "输出到文件";
            this.checkBoxExcel.UseVisualStyleBackColor = true;
            // 
            // printDialogDgv
            // 
            this.printDialogDgv.UseEXDialog = true;
            // 
            // printPreviewDialog1
            // 
            this.printPreviewDialog1.AutoScrollMargin = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.AutoScrollMinSize = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.ClientSize = new System.Drawing.Size(400, 300);
            this.printPreviewDialog1.Enabled = true;
            this.printPreviewDialog1.Icon = ((System.Drawing.Icon)(resources.GetObject("printPreviewDialog1.Icon")));
            this.printPreviewDialog1.Name = "printPreviewDialog1";
            this.printPreviewDialog1.Visible = false;
            // 
            // chkWarn
            // 
            this.chkWarn.AutoSize = true;
            this.chkWarn.CheckAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.chkWarn.Checked = true;
            this.chkWarn.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkWarn.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.chkWarn.Location = new System.Drawing.Point(187, 93);
            this.chkWarn.Name = "chkWarn";
            this.chkWarn.Size = new System.Drawing.Size(78, 17);
            this.chkWarn.TabIndex = 24;
            this.chkWarn.Text = "警告显示";
            this.chkWarn.UseVisualStyleBackColor = true;
            // 
            // checkBoxXH
            // 
            this.checkBoxXH.AutoSize = true;
            this.checkBoxXH.CheckAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.checkBoxXH.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.checkBoxXH.Location = new System.Drawing.Point(262, 94);
            this.checkBoxXH.Name = "checkBoxXH";
            this.checkBoxXH.Size = new System.Drawing.Size(78, 17);
            this.checkBoxXH.TabIndex = 25;
            this.checkBoxXH.Text = "打印序号";
            this.checkBoxXH.UseVisualStyleBackColor = true;
            // 
            // PrintOptions
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(369, 259);
            this.Controls.Add(this.checkBoxXH);
            this.Controls.Add(this.chkWarn);
            this.Controls.Add(this.checkBoxExcel);
            this.Controls.Add(this.checkBoxPrv);
            this.Controls.Add(this.chkFitToPageWidth);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.txtTitle);
            this.Controls.Add(this.gboxRowsToPrint);
            this.Controls.Add(this.lblColumnsToPrint);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.chklst);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "PrintOptions";
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "打印选项";
            this.Load += new System.EventHandler(this.PrintOtions_Load);
            this.gboxRowsToPrint.ResumeLayout(false);
            this.gboxRowsToPrint.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.RadioButton rdoSelectedRows;
        internal System.Windows.Forms.RadioButton rdoAllRows;
        internal System.Windows.Forms.CheckBox chkFitToPageWidth;
        internal System.Windows.Forms.Label lblTitle;
        internal System.Windows.Forms.TextBox txtTitle;
        internal System.Windows.Forms.GroupBox gboxRowsToPrint;
        internal System.Windows.Forms.Label lblColumnsToPrint;
        protected System.Windows.Forms.Button btnOK;
        protected System.Windows.Forms.Button btnCancel;
        internal System.Windows.Forms.CheckedListBox chklst;
        private System.Windows.Forms.CheckBox checkBoxPrv;
        private System.Windows.Forms.CheckBox checkBoxExcel;
        private System.Windows.Forms.PrintDialog printDialogDgv;
        private System.Windows.Forms.PrintPreviewDialog printPreviewDialog1;
        internal System.Windows.Forms.CheckBox chkWarn;
        internal System.Windows.Forms.CheckBox checkBoxXH;

    }
}