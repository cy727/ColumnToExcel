using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using System.Collections;
using System.Data;
using System.Text;
using System.Diagnostics;
using System.IO;
using Microsoft.Office.Interop.Excel;
//using Microsoft.Office;

namespace ColumnToExcel
{
    class PrintDGV
    {
        private static StringFormat StrFormat;  // Holds content of a TextBox Cell to write by DrawString
        private static StringFormat StrFormatR;
        private static StringFormat StrFormatC;  
        private static StringFormat StrFormatComboBox; // Holds content of a Boolean Cell to write by DrawImage
        private static System.Windows.Forms.Button CellButton;       // Holds the Contents of Button Cell
        private static System.Windows.Forms.CheckBox CellCheckBox;   // Holds the Contents of CheckBox Cell 
        private static ComboBox CellComboBox;   // Holds the Contents of ComboBox Cell

        private static int TotalWidth;          // Summation of Columns widths
        private static int RowPos;              // Position of currently printing row 
        private static bool NewPage;            // Indicates if a new page reached
        private static int PageNo;              // Number of pages to print
        private static ArrayList ColumnLefts = new ArrayList();  // Left Coordinate of Columns
        private static ArrayList ColumnWidths = new ArrayList(); // Width of Columns
        private static ArrayList ColumnTypes = new ArrayList();  // DataType of Columns
        private static int CellHeight;          // Height of DataGrid Cell
        private static int RowsPerPage;         // Number of Rows per Page
        private static System.Drawing.Printing.PrintDocument printDoc = new System.Drawing.Printing.PrintDocument();  // PrintDocumnet Object used for printing

        private static string PrintTitle = "";  // Header of pages
        private static DataGridView dgv;        // Holds DataGridView Object to print its contents
        private static List<string> SelectedColumns = new List<string>();   // The Columns Selected by user to print.
        private static List<string> AvailableColumns = new List<string>();  // All Columns avaiable in DataGrid 
        private static bool PrintAllRows = true;   // True = print all rows,  False = print selected rows    
        private static bool PrintPrv=true;    //打印预览
        private static bool PrintWarn = true;  //打印警告
        private static bool PrintXH = false;  //打印警告

        private static bool FitToPageWidth = true; // True = Fits selected columns to page width ,  False = Print columns as showed   
        private static bool FitToFile = true; // Print toFile 
        private static int HeaderHeight = 0;
        private static int WidthXH = 60;
        private static int TopXH=0;
        private static int LeftXH=0;

        private static string FullFileName = string.Empty;

        private static System.Drawing.Font _Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
        private static System.Drawing.Font _Font12 = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));

        public static void Print_DataGridView(DataGridView dgv1,string strTitle,bool bPrv)
        {
            //PrintPreviewDialog ppvw;
            try 
	        {	
                // Getting DataGridView object to print
                dgv = dgv1;

                // Getting all Coulmns Names in the DataGridView
                AvailableColumns.Clear();
                foreach (DataGridViewColumn c in dgv.Columns)
                {
                    if (!c.Visible) continue;
                    AvailableColumns.Add(c.HeaderText);
                }

                // Showing the PrintOption Form
                PrintOptions dlg = new PrintOptions(AvailableColumns, strTitle,bPrv);
                if (dlg.ShowDialog() != DialogResult.OK) return;

                PrintTitle = dlg.PrintTitle;
                PrintAllRows = dlg.PrintAllRows;
                PrintPrv=dlg.PrintPreView;
                PrintWarn = dlg.PrintWarn;
                PrintXH = dlg.PrintXH;

                FitToPageWidth = dlg.FitToPageWidth;
                SelectedColumns = dlg.GetSelectedColumns();
                FitToFile = dlg.PrintToFile;

                if (FitToFile) //输出到EXCEL
                {
                    dlg.TopMost = false;
                    ToExcel(dgv, strTitle);

                    return;
                }

                PrintDialog ppd = new PrintDialog();
                ppd.Document = printDoc;
                if (ppd.ShowDialog() != DialogResult.OK)
                    return;


                RowsPerPage = 0;

                PrintPreviewDialog ppvw = new PrintPreviewDialog();

                ppvw.ShowIcon = false;

                ppvw.Width = 700;
                ppvw.Height =600;
                ppvw.UseAntiAlias = true;

                ppvw.Document = printDoc;

                // Showing the Print Preview Page
                printDoc.BeginPrint += new System.Drawing.Printing.PrintEventHandler(PrintDoc_BeginPrint);
                printDoc.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(PrintDoc_PrintPage);
                if (PrintPrv)
                {
                    if (ppvw.ShowDialog() != DialogResult.OK)
                    {
                        printDoc.BeginPrint -= new System.Drawing.Printing.PrintEventHandler(PrintDoc_BeginPrint);
                        printDoc.PrintPage -= new System.Drawing.Printing.PrintPageEventHandler(PrintDoc_PrintPage);
                        return;
                    }
                }

                // Printing the Documnet
                printDoc.Print();
                printDoc.BeginPrint -= new System.Drawing.Printing.PrintEventHandler(PrintDoc_BeginPrint);
                printDoc.PrintPage -= new System.Drawing.Printing.PrintPageEventHandler(PrintDoc_PrintPage);
	        }
	        catch (Exception ex)
	        {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);        		
	        }
            finally
            {

            }
        }

        private static void PrintDoc_BeginPrint(object sender, 
                    System.Drawing.Printing.PrintEventArgs e) 
        {
            try
	        {
                // Formatting the Content of Text Cell to print
                StrFormat = new StringFormat();
                StrFormat.Alignment = StringAlignment.Near;
                StrFormat.LineAlignment = StringAlignment.Center;
                StrFormat.Trimming = StringTrimming.EllipsisCharacter;

                StrFormatR = new StringFormat();
                StrFormatR.Alignment = StringAlignment.Far;
                StrFormatR.LineAlignment = StringAlignment.Center;
                StrFormatR.Trimming = StringTrimming.EllipsisCharacter;

                StrFormatC = new StringFormat();
                StrFormatC.Alignment = StringAlignment.Center;
                StrFormatC.LineAlignment = StringAlignment.Center;
                StrFormatC.Trimming = StringTrimming.EllipsisCharacter;

                // Formatting the Content of Combo Cells to print
                StrFormatComboBox = new StringFormat();
                StrFormatComboBox.LineAlignment = StringAlignment.Center;
                StrFormatComboBox.FormatFlags = StringFormatFlags.NoWrap;
                StrFormatComboBox.Trimming = StringTrimming.EllipsisCharacter;

                ColumnLefts.Clear();
                ColumnWidths.Clear();
                ColumnTypes.Clear();
                CellHeight = 0;
                RowsPerPage = 0;

                // For various column types
                CellButton = new System.Windows.Forms.Button();
                CellCheckBox = new System.Windows.Forms.CheckBox();
                CellComboBox = new ComboBox();

                // Calculating Total Widths
                TotalWidth = 0;
                foreach (DataGridViewColumn GridCol in dgv.Columns)
                {
                    if (!GridCol.Visible) continue;
                    if (!PrintDGV.SelectedColumns.Contains(GridCol.HeaderText)) continue;
                    TotalWidth += GridCol.Width;
                }
                PageNo = 1;
                NewPage = true;
                RowPos = 0;        		
	        }
	        catch (Exception ex)
	        {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);        		
	        }
        }

        private static void PrintDoc_PrintPage(object sender, 
                    System.Drawing.Printing.PrintPageEventArgs e) 
        {
            int tmpWidth, i, j;
            int eMarginBoundsTop=e.MarginBounds.Top;
            int eMarginBoundsLeft = e.MarginBounds.Left;
            int eMarginBoundsWidth=e.MarginBounds.Width;
            int tmpTop = e.MarginBounds.Top;
            int tmpLeft = e.MarginBounds.Left;
            int itemp;
            string sTemp="";

            try 
	        {	        
                // Before starting first page, it saves Width & Height of Headers and CoulmnType
                if (PageNo == 1) 
                {
                    if (PrintXH) //有序号
                    {
                        eMarginBoundsLeft += WidthXH;
                        eMarginBoundsWidth -= WidthXH;
                        TopXH = e.MarginBounds.Top;
                        LeftXH = e.MarginBounds.Left;
                        tmpLeft += WidthXH;

                    }
                    foreach (DataGridViewColumn GridCol in dgv.Columns)
                    {
                        if (!GridCol.Visible) continue;
                        // Skip if the current column not selected
                        if (!PrintDGV.SelectedColumns.Contains(GridCol.HeaderText)) continue;

                        // Detemining whether the columns are fitted to page or not.
                        if (FitToPageWidth) 
                            tmpWidth = (int)(Math.Floor((double)((double)GridCol.Width / 
                                       (double)TotalWidth * (double)TotalWidth * 
                                       ((double)eMarginBoundsWidth / (double)TotalWidth))));
                        else
                            tmpWidth = GridCol.Width;

                        HeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText,
                                    GridCol.InheritedStyle.Font, tmpWidth).Height) + 11;
                        
                        // Save width & height of headres and ColumnType
                        ColumnLefts.Add(tmpLeft);
                        ColumnWidths.Add(tmpWidth);
                        ColumnTypes.Add(GridCol.GetType());
                        tmpLeft += tmpWidth;
                    }
                }

                // Printing Current Page, Row by Row
                while (RowPos <= dgv.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = dgv.Rows[RowPos];
                    if (GridRow.IsNewRow || (!PrintAllRows && !GridRow.Selected))
                    {
                        RowPos++;
                        continue;
                    }

                    //CellHeight = GridRow.Height;
                    CellHeight = GridRow.Height + 8;

                    if (tmpTop + CellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        DrawFooter(e, RowsPerPage);
                        NewPage = true;
                        PageNo++;
                        e.HasMorePages = true;
                        return;
                    }
                    else
                    {
                        if (NewPage)
                        {
                            // 单据汇总,表头
                            string[] sTitle = PrintTitle.Split(new char[1] { ';'});
                            
                            for (j = 0; j < sTitle.Length;j++ )
                            {
                                if (j <= 1)
                                {
                                    if(j==0)
                                        e.Graphics.DrawString(sTitle[j], _Font12, Brushes.Black, new System.Drawing.RectangleF(e.MarginBounds.Left, tmpTop, e.MarginBounds.Width, (int)(e.Graphics.MeasureString(sTitle[j], _Font12, e.MarginBounds.Width).Height)), StrFormat);
                                    else
                                        e.Graphics.DrawString(sTitle[j], _Font, Brushes.Black, new System.Drawing.RectangleF(e.MarginBounds.Left, tmpTop, e.MarginBounds.Width, (int)(e.Graphics.MeasureString(sTitle[j], _Font12, e.MarginBounds.Width).Height)), StrFormatR);
                                }
                                else
                                    e.Graphics.DrawString(sTitle[j], _Font, Brushes.Black, e.MarginBounds.Left, tmpTop);

                                if (j == 0) continue;

                                if(j==1)
                                    tmpTop += (int)(e.Graphics.MeasureString(sTitle[j], _Font12, e.MarginBounds.Width).Height);
                                else
                                    tmpTop += (int)(e.Graphics.MeasureString(sTitle[j], _Font, e.MarginBounds.Width).Height);

                            }
                            /*
                            e.Graphics.DrawString(PrintTitle, new System.Drawing.Font(dgv.Font, FontStyle.Bold), 
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                            e.Graphics.MeasureString(PrintTitle, new System.Drawing.Font(dgv.Font, 
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            String s = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString();

                            //e.Graphics.DrawString(s, new Font(dgv.Font, FontStyle.Bold), 
                            //        Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width - 
                            //        e.Graphics.MeasureString(s, new Font(dgv.Font, 
                            //        FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top - 
                            //        e.Graphics.MeasureString(PrintTitle, new Font(new Font(dgv.Font, 
                            //        FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            e.Graphics.DrawString(s, new System.Drawing.Font(_Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                                    e.Graphics.MeasureString(s, new System.Drawing.Font(_Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                                    e.Graphics.MeasureString(PrintTitle, new System.Drawing.Font(new System.Drawing.Font(_Font,
                                    FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);
                            */

                            // Draw Columns
                            //tmpTop = e.MarginBounds.Top;
                            i = 0;

                            //表头
                            if (PrintXH) //有序号
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new System.Drawing.Rectangle(LeftXH, tmpTop,
                                    WidthXH, HeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new System.Drawing.Rectangle(LeftXH, tmpTop,
                                    WidthXH, HeaderHeight));

                                e.Graphics.DrawString("序号", _Font,
                                    new SolidBrush(Color.Black),
                                    new System.Drawing.RectangleF(LeftXH, tmpTop,
                                    WidthXH, HeaderHeight), StrFormatC);
                            }
                            foreach (DataGridViewColumn GridCol in dgv.Columns)
                            {
                                if (!GridCol.Visible) continue;
                                if (!PrintDGV.SelectedColumns.Contains(GridCol.HeaderText)) 
                                    continue;

                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new System.Drawing.Rectangle((int)ColumnLefts[i], tmpTop,
                                    (int)ColumnWidths[i], HeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new System.Drawing.Rectangle((int)ColumnLefts[i], tmpTop,
                                    (int)ColumnWidths[i], HeaderHeight));

                                e.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font, 
                                    new SolidBrush(GridCol.InheritedStyle.ForeColor),
                                    new System.Drawing.RectangleF((int)ColumnLefts[i], tmpTop, 
                                    (int)ColumnWidths[i], HeaderHeight), StrFormatC);
                                i++;
                            }
                            NewPage = false;
                            tmpTop += HeaderHeight;
                        }

                        // Draw Columns Contents
                        if (PrintXH) //有序号
                        {
                            e.Graphics.DrawString((RowPos+1).ToString(), _Font, new SolidBrush(Color.Black), new RectangleF(LeftXH, (float)tmpTop, WidthXH, (float)CellHeight), StrFormatC);
                            // Drawing Cells Borders 
                            e.Graphics.DrawRectangle(Pens.Black, new System.Drawing.Rectangle(LeftXH, tmpTop, WidthXH, CellHeight));
                        }
                        i = 0;
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            if (!Cel.OwningColumn.Visible) continue;
                            if (!SelectedColumns.Contains(Cel.OwningColumn.HeaderText))
                                continue;

                            // For the TextBox Column
                            if (((Type) ColumnTypes[i]).Name == "DataGridViewTextBoxColumn" || 
                                ((Type) ColumnTypes[i]).Name == "DataGridViewLinkColumn")
                            {
                                if (PrintWarn)
                                {

                                    //e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font, 
                                    //        new SolidBrush(Cel.InheritedStyle.ForeColor),
                                    //        new RectangleF((int)ColumnLefts[i], (decimal)tmpTop,
                                    //        (int)ColumnWidths[i], (decimal)CellHeight), StrFormat);

                                    //确定背景
                                    if (Cel.Style.BackColor == Color.LightPink)
                                    {
                                        e.Graphics.FillRectangle(new SolidBrush(Color.LightPink), new System.Drawing.Rectangle((int)ColumnLefts[i], (int)tmpTop,(int)ColumnWidths[i], (int)CellHeight));
                                    }
                                }

                                if (Cel.Value.GetType() == typeof(System.DateTime))
                                {
                                    if (Cel.Value.ToString() != "")
                                        sTemp = Convert.ToDateTime(Cel.Value.ToString()).ToString("yyyy年M月dd日");
                                    else
                                        sTemp = "";

                                    e.Graphics.DrawString(sTemp, _Font, new SolidBrush(Cel.InheritedStyle.ForeColor), new RectangleF((int)ColumnLefts[i], (float)tmpTop, (int)ColumnWidths[i], (float)CellHeight), StrFormatC);

                                }
                                else
                                    e.Graphics.DrawString(Cel.Value.ToString(), _Font, 
                                        new SolidBrush(Cel.InheritedStyle.ForeColor),
                                        new RectangleF((int)ColumnLefts[i], (float)tmpTop,
                                        (int)ColumnWidths[i], (float)CellHeight), StrFormatC);

                            }
                            // For the Button Column
                            else if (((Type) ColumnTypes[i]).Name == "DataGridViewButtonColumn")
                            {
                                CellButton.Text = Cel.Value.ToString();
                                CellButton.Size = new Size((int)ColumnWidths[i], CellHeight);
                                Bitmap bmp =new Bitmap(CellButton.Width, CellButton.Height);
                                CellButton.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, 
                                        bmp.Width, bmp.Height));
                                e.Graphics.DrawImage(bmp, new System.Drawing.Point((int)ColumnLefts[i], tmpTop));
                            }
                            // For the CheckBox Column
                            else if (((Type) ColumnTypes[i]).Name == "DataGridViewCheckBoxColumn")
                            {
                                CellCheckBox.Size = new Size(14, 14);
                                CellCheckBox.Checked = (bool)Cel.Value;
                                Bitmap bmp = new Bitmap((int)ColumnWidths[i], CellHeight);
                                Graphics tmpGraphics = Graphics.FromImage(bmp);
                                tmpGraphics.FillRectangle(Brushes.White, new System.Drawing.Rectangle(0, 0, 
                                        bmp.Width, bmp.Height));
                                CellCheckBox.DrawToBitmap(bmp,
                                        new System.Drawing.Rectangle((int)((bmp.Width - CellCheckBox.Width) / 2), 
                                        (int)((bmp.Height - CellCheckBox.Height) / 2), 
                                        CellCheckBox.Width, CellCheckBox.Height));
                                e.Graphics.DrawImage(bmp, new System.Drawing.Point((int)ColumnLefts[i], tmpTop));
                            }
                            // For the ComboBox Column
                            else if (((Type) ColumnTypes[i]).Name == "DataGridViewComboBoxColumn")
                            {
                                CellComboBox.Size = new Size((int)ColumnWidths[i], CellHeight);
                                Bitmap bmp = new Bitmap(CellComboBox.Width, CellComboBox.Height);
                                CellComboBox.DrawToBitmap(bmp, new System.Drawing.Rectangle(0, 0, 
                                        bmp.Width, bmp.Height));
                                e.Graphics.DrawImage(bmp, new System.Drawing.Point((int)ColumnLefts[i], tmpTop));
                                e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font, 
                                        new SolidBrush(Cel.InheritedStyle.ForeColor), 
                                        new RectangleF((int)ColumnLefts[i] + 1, tmpTop, (int)ColumnWidths[i]
                                        - 16, CellHeight), StrFormatComboBox);
                            }
                            // For the Image Column
                            else if (((Type) ColumnTypes[i]).Name == "DataGridViewImageColumn")
                            {
                                System.Drawing.Rectangle CelSize = new System.Drawing.Rectangle((int)ColumnLefts[i], 
                                        tmpTop, (int)ColumnWidths[i], CellHeight);
                                Size ImgSize = ((Image)(Cel.FormattedValue)).Size;
                                e.Graphics.DrawImage((Image)Cel.FormattedValue,
                                        new System.Drawing.Rectangle((int)ColumnLefts[i] + (int)((CelSize.Width - ImgSize.Width) / 2), 
                                        tmpTop + (int)((CelSize.Height - ImgSize.Height) / 2), 
                                        ((Image)(Cel.FormattedValue)).Width, ((Image)(Cel.FormattedValue)).Height));

                            }

                            // Drawing Cells Borders 
                            e.Graphics.DrawRectangle(Pens.Black, new System.Drawing.Rectangle((int)ColumnLefts[i], 
                                    tmpTop, (int)ColumnWidths[i], CellHeight));

                            i++;

                        }
                        tmpTop += CellHeight;
                    }

                    RowPos++;
                    // For the first page it calculates Rows per Page
                    if (PageNo == 1) RowsPerPage++;
                }

                if (RowsPerPage == 0) return;

                // Write Footer (Page Number)
                DrawFooter(e, RowsPerPage);

                e.HasMorePages = false;
	        }
	        catch (Exception ex)
	        {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);        		
	        }
        }

        private static void DrawFooter(System.Drawing.Printing.PrintPageEventArgs e, 
                    int RowsPerPage)
        {
            double cnt = 0; 

            // Detemining rows number to print
            if (PrintAllRows)
            {
                if (dgv.Rows[dgv.Rows.Count - 1].IsNewRow) 
                    cnt = dgv.Rows.Count - 2; // When the DataGridView doesn't allow adding rows
                else
                    cnt = dgv.Rows.Count - 1; // When the DataGridView allows adding rows
            }
            else
                cnt = dgv.SelectedRows.Count;

            if (cnt == 0.0)
                cnt = 1.0;

            // Writing the Page Number on the Bottom of Page
            string PageNum = "第 "+PageNo.ToString() + " 页 共 " + 
                Math.Ceiling((double)(cnt / RowsPerPage)).ToString()+" 页";

            e.Graphics.DrawString(PageNum, dgv.Font, Brushes.Black, 
                e.MarginBounds.Left + (e.MarginBounds.Width - 
                e.Graphics.MeasureString(PageNum, dgv.Font, 
                e.MarginBounds.Width).Width) / 2, e.MarginBounds.Top + 
                e.MarginBounds.Height + 31);
        }

        #region 保存对话框
        private static bool SaveFileDialog()
        {
            SaveFileDialog saveFileDialogOutput = new SaveFileDialog();
            saveFileDialogOutput.Filter = "excel files(*.xls)|*.xls";//excel files(*.xls)|*.xls|All files(*.*)|*.*
            saveFileDialogOutput.FilterIndex = 0;
            saveFileDialogOutput.RestoreDirectory = true;
            saveFileDialogOutput.CreatePrompt = true;

            if (saveFileDialogOutput.ShowDialog() != DialogResult.OK) return false;

            FullFileName=saveFileDialogOutput.FileName.ToString();

            return true;
        }
        #endregion


        private static void ToExcel(DataGridView ExportGrid, string p_ReportName)
        {
            //如果网格尚未数据绑定
            if(ExportGrid==null)
                return;

            // 列索引，行索引
            int colIndex = 0;
            int rowIndex = 0;
            int j;

            //总可见列数，总可见行数
            //int colCount = ExportGrid.Columns.GetColumnCount(DataGridViewElementStates.Visible);
            //int rowCount = ExportGrid.Rows.GetRowCount(DataGridViewElementStates.Visible);
            int rowCount = 0;
            int colCount=SelectedColumns.Count;
            if (PrintAllRows) //打印全部
            {
                rowCount = ExportGrid.Rows.GetRowCount(DataGridViewElementStates.Visible);
            }
            else //打印选择
            {
                rowCount = ExportGrid.Rows.GetRowCount(DataGridViewElementStates.Selected);
            }

            //如果DataGridView中没有行，返回
            if (rowCount == 0)
                return;

            //保存对话框
            if (!SaveFileDialog())
                return;

            // 创建Excel对象                    
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            if (xlApp == null)
            {
                MessageBox.Show("Excel无法启动","系统信息");
                return;
            }
            // 创建Excel工作薄
            Microsoft.Office.Interop.Excel.Workbook xlBook = xlApp.Workbooks.Add(true);
            Microsoft.Office.Interop.Excel.Worksheet xlSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlBook.Worksheets[1];

            // 设置标题，实测中发现执行设置字体大小和将字体设置为粗体的语句耗时较长，故注释掉了
            string[] sTitle = p_ReportName.Split(new char[1] { ';' });
            Microsoft.Office.Interop.Excel.Range range = xlSheet.get_Range(xlApp.Cells[1, 1], xlApp.Cells[1, colCount]);
            range.MergeCells = true;
            xlApp.ActiveCell.FormulaR1C1 = sTitle[0];
            //xlApp.ActiveCell.Font.Size = 20;
            //xlApp.ActiveCell.Font.Bold = true;
            xlApp.ActiveCell.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

            // 单据汇总


            // 创建缓存数据
            object[,] objData = new object[rowCount + 2 + sTitle.Length, colCount];
            for (j = 1; j < sTitle.Length; j++)
            {
                //Microsoft.Office.Interop.Excel.Range range1 = xlSheet.get_Range(xlApp.Cells[rowIndex, 1], xlApp.Cells[rowIndex, colCount]);
                //range1.MergeCells = true;
                objData[rowIndex, 0] = sTitle[j];
                rowIndex++;
            }
            objData[rowIndex, 0] = "";
            rowIndex++;




            // 获取列标题，隐藏的列不处理
            for (int i = 0; i < ExportGrid.ColumnCount; i++)
            {
                if (!PrintDGV.SelectedColumns.Contains(ExportGrid.Columns[i].HeaderText)) continue;
                if (ExportGrid.Columns[i].Visible)
                    objData[rowIndex, colIndex++] = ExportGrid.Columns[i].HeaderText;
            }
            // 获取数据，隐藏的列的数据忽略

            for (int i = 1; i <= ExportGrid.Rows.GetRowCount(DataGridViewElementStates.Visible); i++)
            {
                if (!PrintAllRows && !ExportGrid.Rows[i-1].Selected)
                {
                    continue;
                }
                rowIndex++;
                colIndex = 0;


                for (j = 0; j < ExportGrid.ColumnCount; j++)
                {
                    if (!PrintDGV.SelectedColumns.Contains(ExportGrid.Columns[j].HeaderText)) continue;
                    if (ExportGrid.Columns[j].Visible)
                    {
                        if (ExportGrid[j, i - 1].Value != null)
                            objData[rowIndex, colIndex++] = ExportGrid[j, i - 1].Value.ToString();
                        else
                            objData[rowIndex, colIndex++] = "";
                    }
                }
                
                System.Windows.Forms.Application.DoEvents();
            }

            // 写入Excel
            //xlApp.get_Range(xlApp.Cells[2, 1], xlApp.Cells[2, colIndex]).Font.Bold = true;
            range = xlSheet.get_Range(xlApp.Cells[2, 1], xlApp.Cells[rowCount + +2 + sTitle.Length, colCount]);
            range.Value2 = objData;

            // 保存
            try
            {
                xlApp.Cells.EntireColumn.AutoFit();
                xlApp.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                xlApp.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                //xlApp.Visible   =   true;   
                xlBook.Saved = true;
                xlBook.SaveCopyAs(FullFileName);
                
                MessageBox.Show("导出成功！","系统信息");
                
                
            }
            catch
            {
                MessageBox.Show("保存出错，请检查文件是否被正使用！", "系统信息");
                //return false;
            }
            finally
            {
                xlApp.Quit();
                GC.Collect();
                //KillProcess("excel");
            }
            //return true;
        }

       #region 杀死进程
       private void KillProcess(string processName)
        {
            System.Diagnostics.Process myproc = new System.Diagnostics.Process();
            //得到所有打开的进程 
            try
            {
                foreach (Process thisproc in Process.GetProcessesByName(processName))
                {
                    thisproc.Kill();
                }
            }
            catch (Exception Exc)
            {
                throw new Exception("", Exc);
            }
        }
        #endregion




    }
}
