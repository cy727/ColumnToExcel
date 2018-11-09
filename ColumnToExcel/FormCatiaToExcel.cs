using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;

using INFITF;
using MECMOD;
using HybridShapeTypeLib;
using PARTITF;
using ProductStructureTypeLib;
using KnowledgewareTypeLib;
using SPATypeLib;
using DRAFTINGITF;

namespace ColumnToExcel
{
    public partial class FormCatiaToExcel : Form
    {

        bool bAllSel = true;

        Selection cSelection;
        private ArrayList cElement = new ArrayList();

        private ArrayList sufElement = new ArrayList();
        private ArrayList pointElement = new ArrayList();

        private int ColumnCount = 0; //数量

        INFITF.Application CATIA;
        INFITF.Documents docCATIA;
        MECMOD.PartDocument partDocCATIA;
        MECMOD.Part partCATIA;
        MECMOD.Bodies bodiesCATIA;
        MECMOD.Body bodyCATIA;
        MECMOD.Sketches skCATIA;

        ClassCATIA cCATIA = new ClassCATIA();
        DataTable dtCATIA = new DataTable();
        DataTable drawingCATIA = new DataTable();


        Workbench SPAworkbench;
        FormXYOption formXYOption = new FormXYOption();
        FormDrawingsOption formDrawingsOption = new FormDrawingsOption();


        const double Pi = 3.14159265358979;
        const double EP = 1e-5;

        public FormCatiaToExcel()
        {
            InitializeComponent();
        }

        private void buttonP_Click(object sender, EventArgs e)
        {
            //格式确认
            this.TopMost = false;
            string strT = "顶点参数";
            PrintDGV.Print_DataGridView(dataGridViewCATIA, strT, true);
        }

        private void FormColumnToExcel_Load(object sender, EventArgs e)
        {
            //this.TopLevel = true;


            dtCATIA.Columns.Add("N", System.Type.GetType("System.String"));
            dtCATIA.Columns.Add("Name", System.Type.GetType("System.String"));
            dtCATIA.Columns.Add("PointName", System.Type.GetType("System.String"));

            dtCATIA.Columns.Add("X", System.Type.GetType("System.Decimal"));
            dtCATIA.Columns.Add("Y", System.Type.GetType("System.Decimal"));
            dtCATIA.Columns.Add("Z", System.Type.GetType("System.Decimal"));
            dtCATIA.Columns.Add("SUFID", System.Type.GetType("System.Decimal"));
            dtCATIA.Columns.Add("POINTID", System.Type.GetType("System.Decimal"));


            dataGridViewCATIA.DataSource = dtCATIA;

            dataGridViewCATIA.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewCATIA.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewCATIA.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewCATIA.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewCATIA.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewCATIA.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewCATIA.Columns[6].Visible = false;
            dataGridViewCATIA.Columns[7].Visible = false;

            try
            {
                CATIA = (INFITF.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("CATIA.Application");
            }
            catch
            {
                System.Type CATIAType = System.Type.GetTypeFromProgID("CATIA.Application");
                CATIA = (INFITF.Application)System.Activator.CreateInstance(CATIAType);

                //return;
            }
            CATIA.Visible = true;
            cCATIA.CATIA = CATIA;
            //this.TopMost = true;

            cCATIA.InitCATIAPart();
            //cCATIA.oPartDoc.GetWorkbench("SPAworkbench");

            drawingCATIA.Columns.Add("ID", System.Type.GetType("System.Decimal")); 
            drawingCATIA.Columns.Add("LOWX", System.Type.GetType("System.Decimal"));//下标
            drawingCATIA.Columns.Add("LOWY", System.Type.GetType("System.Decimal"));
            drawingCATIA.Columns.Add("HIGHX", System.Type.GetType("System.Decimal"));//上标
            drawingCATIA.Columns.Add("HIGHY", System.Type.GetType("System.Decimal"));
            drawingCATIA.Columns.Add("OriX", System.Type.GetType("System.Decimal"));//原点（中点）
            drawingCATIA.Columns.Add("OriY", System.Type.GetType("System.Decimal"));


        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonAll_Click(object sender, EventArgs e)
        {
            cCATIA.InitCATIAPart();
            bAllSel = true;
            ColumnCount = 0;
            cElement.Clear();

            int i, j;

            for (i = 1; i <= cCATIA.oPartDoc.Part.HybridBodies.Count; i++)
            {
                HybridBody myBody = cCATIA.oPart.HybridBodies.Item(i);

                if (myBody.get_Name().IndexOf(textBoxZMC.Text.Trim()) != -1) //???????????????
                {

                    cElement.Add((HybridBody)myBody);
                }

            }

            ColumnCount = cElement.Count;
            toolStripStatusLabelCount.Text = ColumnCount.ToString();
        }

        private void buttonSel_Click(object sender, EventArgs e)
        {
            int i;
            bAllSel = false;
            this.TopMost = true;

            ColumnCount = 0;
            cElement.Clear();

            cCATIA.InitCATIAPart();
            ArrayList InputObject = new ArrayList();
            cSelection = cCATIA.oPartDoc.Selection;


            //InputObject.Add("AnyObject");
            //InputObject.Add("BiDim");
            //InputObject.Add("HybridShapeSurfaceExplicit");
            InputObject.Add("HybridBody");
            


            string tt = "请选择元素（ESC退出,ENTER确定）";
            string Status;

            Status = cSelection.SelectElement3(InputObject.ToArray(), ref tt, true, CATMultiSelectionMode.CATMultiSelTriggWhenSelPerf, false);

            if (Status == "Normal")
            {
                ColumnCount = cSelection.Count;
                for (i = 1; i <= cSelection.Count; i++)
                {
                    if (((HybridBody)(cSelection.Item(i).Value)).get_Name().IndexOf(textBoxZMC.Text.Trim()) != -1) //???????????????
                    {

                            cElement.Add((HybridBody)cSelection.Item(i).Value);
                    }
                }
            }
            toolStripStatusLabelCount.Text = cElement.Count.ToString();

        }

        private void buttonJG_Click(object sender, EventArgs e)
        {
            if (cElement.Count < 1)
            {
                MessageBox.Show("请先选定统计元素", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            


            int i, j, k, no;
            Workbench TheSPAWorkbench;
            TheSPAWorkbench = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench");

            Measurable TheMeasurable;

            dtCATIA.Clear();
            pointElement.Clear();
            sufElement.Clear();

            //Selection selS = cCATIA.oPartDoc.Selection;
            Selection selS = cCATIA.oPartDoc.Selection.Selection;
            bool bAdd=false;
            no = 1;


            for (i = 0; i < cElement.Count; i++)
            {
                    HybridBody hybHB;

                    hybHB=(HybridBody)cElement[i];

                    toolStripStatusLabelCATIA.Text = "正在导出顶点....";
                    toolStripProgressBarCATIA.Maximum = hybHB.HybridShapes.Count;


                    for(j=1;j<=hybHB.HybridShapes.Count; j++)
                    {
                        try
                        {
                            HybridShape hbsSHAPE = (HybridShape)hybHB.HybridShapes.Item(j);
                            

                            if(hbsSHAPE.get_Name().IndexOf(textBoxMMC.Text.Trim()) != -1)
                            {
                                selS.Clear();

                                selS.Add((AnyObject)hbsSHAPE);
                                selS.Search("((((((CATStFreeStyleSearch.Point + CAT2DLSearch.2DPoint) + CATSketchSearch.2DPoint) + CATDrwSearch.2DPoint) + CATPrtSearch.Point) + CATGmoSearch.Point) + CATSpdSearch.Point),sel");
                                bAdd = false;

                                for (k = 0; k < selS.Count; k++)
                                {
                                    AnyObject hyPoint = (AnyObject)(selS.Item(k + 1).Value);
                                    if (hyPoint.get_Name().IndexOf(textBoxDMC.Text.Trim()) == -1)
                                        continue;

                                    if (!bAdd)
                                    {
                                        sufElement.Add(hbsSHAPE);
                                        bAdd = true;
                                    }
                                    pointElement.Add(hyPoint);

                                    Reference refP = cCATIA.oPart.CreateReferenceFromObject(hyPoint);

                                    TheMeasurable = ((SPAWorkbench)TheSPAWorkbench).GetMeasurable(refP);

                                    

                                    object[] oPoint = new object[3];

                                    TheMeasurable.GetPoint(oPoint);

                                    //hyP.GetCoordinates((Array)oPoint);

                                    object[] oTemp = new object[8];

                                    oTemp[0] = no.ToString();
                                    oTemp[1] = hbsSHAPE.get_Name();
                                    oTemp[2] = hyPoint.get_Name();
                                    oTemp[3] = oPoint[0].ToString();
                                    oTemp[4] = oPoint[1].ToString();
                                    oTemp[5] = oPoint[2].ToString();
                                    oTemp[6] = sufElement.Count-1;
                                    oTemp[7] = pointElement.Count - 1;


                                    no++;



                                    dtCATIA.Rows.Add(oTemp);
                                   
                                }

                           }
                        }
                        catch
                        {
                            continue;
                        }
                        toolStripProgressBarCATIA.Value = j;
                    }
                    


            }
            MessageBox.Show("顶点导出完毕，请选择打印或输出到EXCEL", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            toolStripStatusLabelCATIA.Text = "顶点导出完毕";
            toolStripProgressBarCATIA.Value = toolStripProgressBarCATIA.Minimum;

        }

        private void buttonGM_Click(object sender, EventArgs e)
        {
            cSelection = cCATIA.oPartDoc.Selection;
            int i,j,k;
            this.TopMost = false;

            k = 0;
            if (cSelection.Count < 1)
            {
                if (MessageBox.Show("是否对所有元素更名？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }

                for (i = 1; i <= cCATIA.oPartDoc.Part.HybridBodies.Count; i++)
                {
                    HybridBody myBody = cCATIA.oPart.HybridBodies.Item(i);

                    if (myBody.get_Name().IndexOf(textBoxZMC.Text.Trim()) != -1) //???????????????
                    {
                        for (j = 1; j <= myBody.HybridShapes.Count; j++)
                        {
                            HybridShape myShape=myBody.HybridShapes.Item(j);
                            if (myShape.get_Name().IndexOf(textBoxMMC.Text.Trim()) != -1)
                            {
                                myShape.set_Name(getStringNum(k));
                                k++;
                            }

                        }
                    }

                }


            }
            else
            {
                for (i = 1; i <= cSelection.Count; i++)
                {
                    HybridBody myBody1;
                    try
                    {
                        myBody1 = (HybridBody)cSelection.Item(i).Value;
                    }
                    catch
                    {
                        continue;
                    }
                    if (myBody1.get_Name().IndexOf(textBoxZMC.Text.Trim()) != -1) //???????????????
                    {
                        for (j = 1; j <= myBody1.HybridShapes.Count; j++)
                        {
                            HybridShape myShape1 = myBody1.HybridShapes.Item(j);
                            if (myShape1.get_Name().IndexOf(textBoxMMC.Text.Trim()) != -1)
                            {
                                myShape1.set_Name(getStringNum(k));
                                k++;
                            }

                        }
                    }

                }
            }
            cCATIA.oPart.Update();
            MessageBox.Show("命名更新完毕");
        }

        private string getStringNum(int iNum)
        {
            return textBoxQZ.Text.Trim()+(numericUpDownQS.Value + iNum * (numericUpDownJG.Value)).ToString().PadLeft((int)numericUpDownZC.Value, '0')+textBoxHZ.Text.Trim();
        }

        private void dataGridViewCATIA_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int i,j;

            if (dataGridViewCATIA.SelectedRows.Count < 1)
                return;
            Selection selS = cCATIA.oPartDoc.Selection.Selection;
            selS.Clear();

            /*
            for (i = 0; i < cElement.Count; i++)
            {

                HybridBody hybHB;

                hybHB = (HybridBody)cElement[i];

                for (j = 1; j <= hybHB.HybridShapes.Count; j++)
                {
                    try
                    {
                        HybridShape hbsSHAPE = (HybridShape)hybHB.HybridShapes.Item(j);
                        if (hbsSHAPE.get_Name() == dataGridViewCATIA.SelectedRows[0].Cells[1].Value.ToString())
                        {
                            selS.Add((AnyObject)hbsSHAPE);
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

            }
             */

            if (dataGridViewCATIA.SelectedRows[0].Cells[6].Value.ToString() == "")
                return;
            if (dataGridViewCATIA.SelectedRows[0].Cells[7].Value.ToString() == "")
                return;
            try
            {
                selS.Add((AnyObject)sufElement[int.Parse(dataGridViewCATIA.SelectedRows[0].Cells[6].Value.ToString())]);
            }
            catch
            {
                return;
            }

            cCATIA.oPart.Update();


        }

        private void 选择点ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int i, j;

            if (dataGridViewCATIA.SelectedRows.Count < 1)
                return;
            Selection selS = cCATIA.oPartDoc.Selection.Selection;
            selS.Clear();

            if (dataGridViewCATIA.SelectedRows[0].Cells[6].Value.ToString() == "")
                return;
            if (dataGridViewCATIA.SelectedRows[0].Cells[7].Value.ToString() == "")
                return;
            try
            {
                selS.Add((AnyObject)pointElement[int.Parse(dataGridViewCATIA.SelectedRows[0].Cells[7].Value.ToString())]);
            }
            catch
            {
                return;
            }

            cCATIA.oPart.Update();
        }

        private void 选择面ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridViewCATIA_CellContentDoubleClick(null,null);
        }

        private ProductDocument oProductDoc;
        private MECMOD.PartDocument oPartDocOut;
        private MECMOD.Part oPartOut;

        private MECMOD.Bodies oBodiesOut;
        private MECMOD.Body oBodyOut;

        private MECMOD.HybridBodies oHBodiesOut;
        private MECMOD.HybridBody oHBodyOut;

        public ShapeFactory oSFOut;
        public HybridShapeFactory oHSFOut;


        private void buttonDC_Click(object sender, EventArgs e)
        {
            int i, j, k1, k2;
            Reference refElement;
            Reference r1, r2, r3;

            formXYOption.ShowDialog();

            if (formXYOption.bCancel)
                return;

            Workbench TheSPAWorkbench;
            Measurable TheMeasurable;


            //初始化
            try
            {
                oPartDocOut = (MECMOD.PartDocument)CATIA.Documents.Add("Part");
                TheSPAWorkbench = oPartDocOut.GetWorkbench("SPAWorkbench");

                oPartOut = oPartDocOut.Part;
                oBodiesOut = oPartOut.Bodies;
                oBodyOut = oPartOut.MainBody;
                oHBodiesOut = oPartOut.HybridBodies;

                oSFOut = (ShapeFactory)oPartOut.ShapeFactory;
                oHSFOut = (HybridShapeFactory)oPartOut.HybridShapeFactory;

            }
            catch
            {
                return;
            }

            MECMOD.HybridBody oHB = oHBodiesOut.Add();
            oHB.set_Name("Original");

            bool bCatiaBOOL;

            //写入原点
            toolStripStatusLabelCATIA.Text = "正在写入原始点....";
            toolStripProgressBarCATIA.Maximum = dataGridViewCATIA.RowCount;
            for (i = 0; i < dataGridViewCATIA.RowCount; i++)
            {
                if(dataGridViewCATIA.Rows[i].IsNewRow) 
                    continue;

                //
                bCatiaBOOL = true;
                MECMOD.HybridBody oHBSurf=null;
                while (true)
                {
                    if (bCatiaBOOL)
                    {
                        oHBSurf = oHB.HybridBodies.Add();
                        oHBSurf.set_Name(dataGridViewCATIA.Rows[i].Cells[1].Value.ToString());
                        bCatiaBOOL = false;
                    }

                    HybridShapePointCoord hspcOut = oHSFOut.AddNewPointCoord(double.Parse(dataGridViewCATIA.Rows[i].Cells[3].Value.ToString()), double.Parse(dataGridViewCATIA.Rows[i].Cells[4].Value.ToString()), double.Parse(dataGridViewCATIA.Rows[i].Cells[5].Value.ToString()));
                    hspcOut.set_Name(dataGridViewCATIA.Rows[i].Cells[2].Value.ToString());
                    if (oHBSurf!=null)
                        oHBSurf.AppendHybridShape(hspcOut);

                    if (i + 1 == dataGridViewCATIA.RowCount)
                        break;

                    if (dataGridViewCATIA.Rows[i+1].Cells[1].Value.ToString() != dataGridViewCATIA.Rows[i].Cells[1].Value.ToString())
                    {
                        break;
                    }

                    i++;

                    if (i == dataGridViewCATIA.RowCount)
                        break;
                }

                toolStripProgressBarCATIA.Value = i + 1;

            }
            toolStripProgressBarCATIA.Value = toolStripProgressBarCATIA.Minimum;

            refElement = oPartOut.CreateReferenceFromObject((INFITF.AnyObject)oHB);
            oHSFOut.GSMVisibility(refElement,0);
            oPartOut.Update();

            //确定轴,确定旋转点
            MECMOD.HybridBody oHB2 = oHBodiesOut.Add();
            oHB2.set_Name("RotateAxis");

            //
            AxisSystems axisSystemsOut = oPartOut.AxisSystems;
            Reference rfT;
           
            MECMOD.HybridBody oHBS;
            Selection sl;
            sl = oPartDocOut.Selection;
            sl.Clear();

            //去除不对面(两点以下)
            for (i = 1; i <= oHB.HybridBodies.Count; i++)
            {
                oHBS = oHB.HybridBodies.Item(i);
                if (oHBS.HybridShapes.Count <= 2) //小于两点无法成面
                {
                    sl.Add(oHBS);
                    //oHB.HybridBodies.
                }
            }
            if(sl.Count>=1)
                sl.Delete();


            toolStripStatusLabelCATIA.Text = "正在导出旋转点....";
            toolStripProgressBarCATIA.Maximum = oHB.HybridBodies.Count;
            for (i = 1; i <= oHB.HybridBodies.Count; i++)
            {
                oHBS = oHB.HybridBodies.Item(i);
                if (oHBS.HybridShapes.Count <= 2) //小于两点无法成面
                    continue;

                MECMOD.HybridBody oHBSurf2 = oHB2.HybridBodies.Add();
                oHBSurf2.set_Name(oHBS.get_Name());

                AxisSystem axisSystemOut = axisSystemsOut.Add();
                axisSystemOut.set_Name("Axis_" + oHBS.get_Name());

                //P1 O
                HybridShapePointCoord hspc1 = (HybridShapePointCoord)oHBS.HybridShapes.Item(1);
                
                
                axisSystemOut.OriginType = CATAxisSystemOriginType.catAxisSystemOriginByPoint;

                rfT = oPartOut.CreateReferenceFromObject(hspc1);
                axisSystemOut.OriginPoint = rfT;

                //P2 X
                axisSystemOut.XAxisType = CATAxisSystemAxisType.catAxisSystemAxisSameDirection;
                HybridShapePointCoord hspc2 = (HybridShapePointCoord)oHBS.HybridShapes.Item(2);
                rfT = oPartOut.CreateReferenceFromObject(hspc2);
                axisSystemOut.XAxisDirection = rfT;


                //P3 Y
                axisSystemOut.YAxisType = CATAxisSystemAxisType.catAxisSystemAxisSameDirection;
                HybridShapePointCoord hspc3 = (HybridShapePointCoord)oHBS.HybridShapes.Item(3);
                rfT = oPartOut.CreateReferenceFromObject(hspc3);
                axisSystemOut.YAxisDirection = rfT;

                //Z轴缺省
                axisSystemOut.ZAxisType = CATAxisSystemAxisType.catAxisSystemAxisSameDirection;
                /*
                HybridShapeFill hsf1 = oHSFOut.AddNewFill();

                //HybridShapeLinePtPt hspp = oHSFOut.AddNewLinePtPt((Reference)hspc1, (Reference)hspc2);
                hsf1.AddBound(oPartOut.CreateReferenceFromObject(oHSFOut.AddNewLinePtPt((Reference)hspc1,(Reference)hspc2)));
                hsf1.AddBound(oPartOut.CreateReferenceFromObject(oHSFOut.AddNewLinePtPt((Reference)hspc2, (Reference)hspc3)));
                hsf1.AddBound(oPartOut.CreateReferenceFromObject(oHSFOut.AddNewLinePtPt((Reference)hspc3, (Reference)hspc1)));
                rfT = oPartOut.CreateReferenceFromObject(hsf1);

                HybridShapeLineNormal hsln1 = oHSFOut.AddNewLineNormal(rfT, (Reference)hspc1, 0, 100, false);
                oHBSurf2.AppendHybridShape(hsln1);

                HybridShapePointOnCurve initPoint = oHSFOut.AddNewPointOnCurveWithReferenceFromDistance((Reference)hsln1, (Reference)hspc1, 50, false);
                oHBSurf2.AppendHybridShape(initPoint);
                 */

                /*
                Reference refP = oPartOut.CreateReferenceFromObject(initPoint);
                TheMeasurable = ((SPAWorkbench)TheSPAWorkbench).GetMeasurable(refP);

                object[] oPoint = new object[3];
                TheMeasurable.GetPoint(oPoint);
                axisSystemOut.PutZAxis(oPoint);
                 */
                oPartOut.UpdateObject(axisSystemOut);

                object[] axisOrigin=new object[3];
                object[] xVect=new object[3];
                object[] yVect=new object[3];
                object[] zVect=new object[3];

                axisSystemOut.GetOrigin(axisOrigin);
                axisSystemOut.GetXAxis(xVect);
                axisSystemOut.GetYAxis(yVect);
                axisSystemOut.GetZAxis(zVect);

                NormalizeVector(xVect, ref xVect);
                NormalizeVector(yVect, ref yVect);
                NormalizeVector(zVect, ref zVect);

                
                object[] oPoint = new object[3];
                object[] globalCoords=new object[3];
                object[] delta = new object[3];
                object[] csCoords = new object[3];

                for (j = 1; j <= oHBS.HybridShapes.Count; j++)
                {
                    HybridShapePointCoord hspcT = (HybridShapePointCoord)oHBS.HybridShapes.Item(j);
                    hspcT.GetCoordinates(globalCoords);

                    delta[0] = (double)globalCoords[0] - (double)axisOrigin[0];
                    delta[1] = (double)globalCoords[1] - (double)axisOrigin[1];
                    delta[2] = (double)globalCoords[2] - (double)axisOrigin[2];

                    csCoords[0] = DotProduct(delta, xVect);
                    csCoords[1] = DotProduct(delta, yVect);
                    csCoords[2] = DotProduct(delta, zVect);

                    //hspcT.RefAxisSystem = rfT;

                    HybridShapePointCoord pXY = oHSFOut.AddNewPointCoord((double)csCoords[0], (double)csCoords[1], (double)csCoords[2]);
                    pXY.set_Name(oHBS.HybridShapes.Item(j).get_Name());
                    oHBSurf2.AppendHybridShape(pXY);
                }

                toolStripProgressBarCATIA.Value = i;

            }
            refElement = oPartOut.CreateReferenceFromObject((INFITF.AnyObject)oHB2);
            oHSFOut.GSMVisibility(refElement, 0);

            toolStripProgressBarCATIA.Value = toolStripProgressBarCATIA.Minimum;


            /*
             * //确定轴,确定旋转点
            Reference RefBasePlane;
            RefBasePlane = oPartOut.CreateReferenceFromObject(oPartOut.OriginElements.PlaneXY);

            
            bool bPlaneXY = true; //是否选择XY作为第二元素
            
            for (i = 1; i <= oHB.HybridBodies.Count; i++)
            {
                oHBS = oHB.HybridBodies.Item(i);
                if (oHBS.HybridShapes.Count <= 1)
                    continue;

                bCatiaBOOL = true;
                k1 = 0;
                k2 = 0;

                for (j = 1; j <= oHBS.HybridShapes.Count; j++)
                {
                    if (j == oHBS.HybridShapes.Count)
                        k2 = 1;
                    else
                        k2 = j + 1;

                    HybridShapePointCoord hspc1 = (HybridShapePointCoord)oHBS.HybridShapes.Item(j);
                    HybridShapePointCoord hspc2 = (HybridShapePointCoord)oHBS.HybridShapes.Item(k2);

                    if (cCATIA.IsSameNum(hspc1.X.Value, hspc2.X.Value) && cCATIA.IsSameNum(hspc1.Y.Value, hspc2.Y.Value)) //投影到一个点
                    {
                        continue;
                    }

                    //两点投影不重合
                    k1 = j;

                    HybridShapePointCoord hspcL1 = oHSFOut.AddNewPointCoord(hspc1.X.Value, hspc1.Y.Value, 0);
                    HybridShapePointCoord hspcL2 = oHSFOut.AddNewPointCoord(hspc2.X.Value, hspc2.Y.Value, 0);

                    r1 = oPartOut.CreateReferenceFromObject(hspcL1);
                    r2 = oPartOut.CreateReferenceFromObject(hspcL2);

                    if (cCATIA.IsSameNum(hspc1.X.Value, hspc2.X.Value)) //X轴相同，选择YZ面
                    {
                        bPlaneXY = false;
                    }
                    else
                    {
                        bPlaneXY = true;
                    }


                    HybridShapeLinePtPt hyPP = oHSFOut.AddNewLinePtPt(r1,r2);
                    hyPP.set_Name("ProjectAxis");
                    oHBS.AppendHybridShape(hyPP);
                    
                    bCatiaBOOL = false;
                    break;
                }

                oPartOut.Update();

                if (bCatiaBOOL)
                    continue;

                //得到旋转点 k1, k1+1(k2)为轴点
                if (k1 == oHBS.HybridShapes.Count)
                    k2 = 1;
                else
                    k2 = k1 + 1;

                MECMOD.HybridBody oHBSurf2 = oHB2.HybridBodies.Add();
                oHBSurf2.set_Name(oHBS.get_Name());

                if(bPlaneXY)
                     RefBasePlane = oPartOut.CreateReferenceFromObject(oPartOut.OriginElements.PlaneZX);
                else
                    RefBasePlane = oPartOut.CreateReferenceFromObject(oPartOut.OriginElements.PlaneYZ);

                for (j = 1; j < oHBS.HybridShapes.Count; j++) //最后一个为线
                {
                    if (j == k1 || j == k2) //不需要
                    {
                        continue;
                    }

                    HybridShapeRotate hsr = oHSFOut.AddNewEmptyRotate();
                    hsr.ElemToRotate = (Reference)(oHBS.HybridShapes.Item(j));
                    hsr.VolumeResult = false;
                    hsr.RotationType = 1;

                    r1 = oPartOut.CreateReferenceFromObject(oHBS.HybridShapes.Item(oHBS.HybridShapes.Count));
                    hsr.Axis = r1;

                    r2 = oPartOut.CreateReferenceFromObject(oHBS.HybridShapes.Item(j));
                    hsr.FirstElement = r2;

                    hsr.SecondElement = RefBasePlane;
                    hsr.OrientationOfFirstElement = false;
                    hsr.OrientationOfSecondElement = false;

                    hsr.set_Name(oHBS.HybridShapes.Item(j).get_Name());
                    oHBSurf2.AppendHybridShape(hsr);

                    //r1=
             


                }
             

                
            }
            */

            //排列点

            MECMOD.HybridBody oHB3 = oHBodiesOut.Add();
            oHB3.set_Name("ResultArray");

            double[] dPointOrgin = new double[3];
            double[] dPoint = new double[3];

            dPointOrgin[0] = (double)formXYOption.numericUpDownX.Value;
            dPointOrgin[1] = (double)formXYOption.numericUpDownY.Value;

            double[] dPointOrginEle = new double[3];
            double dMaxX = 0, dMaxXLength = 0;
            double dMinY = 0, dMinYLength = 0, dYLength = 0 ;

            int iNumOfArray = (int)formXYOption.numericUpDownLS.Value; //列数

            if(formXYOption.checkBoxBFL.Checked) //不分列
                iNumOfArray = oHB2.HybridBodies.Count;

            ArrayList alPoint = new ArrayList();
            ArrayList alLine = new ArrayList();

            toolStripStatusLabelCATIA.Text = "正在生成最终结果....";
            toolStripProgressBarCATIA.Maximum = oHB2.HybridBodies.Count;

            k1 = 1; //列
            for (i = 1; i <= oHB2.HybridBodies.Count; i++)
            {
                oHBS = oHB2.HybridBodies.Item(i);
                if (oHBS.HybridShapes.Count <= 2) //小于两点无法成面,有一个定义坐标
                    continue;

                if (k1 > iNumOfArray) //换行
                {
                    dPointOrgin[0] = (double)formXYOption.numericUpDownX.Value;
                    dPointOrgin[1] -= dYLength + (double)formXYOption.numericUpDownLJJ.Value;
                    dYLength = 0; dMaxXLength = 0;dMinYLength = 0;
                    k1 = 1;
                }

                MECMOD.HybridBody oHBSurf3 = oHB3.HybridBodies.Add();
                oHBSurf3.set_Name(oHBS.get_Name());

                //得到原点（左上角）
                HybridShapePointCoord hspcEle = (HybridShapePointCoord)oHBS.HybridShapes.Item(1);

                dPointOrginEle[0] = hspcEle.X.Value;
                dPointOrginEle[1] = hspcEle.Y.Value;

                dMaxX = hspcEle.X.Value;
                dMinY = hspcEle.Y.Value;

                for (j = 2; j <= oHBS.HybridShapes.Count; j++)
                {
                    hspcEle = (HybridShapePointCoord)oHBS.HybridShapes.Item(j);
                    if (hspcEle.X.Value < dPointOrginEle[0])
                    {
                        dPointOrginEle[0] = hspcEle.X.Value;
                    }
                    if (hspcEle.Y.Value > dPointOrginEle[1])
                    {
                        dPointOrginEle[1] = hspcEle.Y.Value;
                    }
                    if (hspcEle.X.Value > dMaxX)
                    {
                        dMaxX = hspcEle.X.Value;
                    }
                    if (hspcEle.Y.Value < dMinY)
                    {
                        dMinY = hspcEle.Y.Value;
                    }

                }

                dMaxXLength = dMaxX - dPointOrginEle[0];
                dMinYLength = dPointOrginEle[1] - dMinY;

                alPoint.Clear();
                for (j = 1; j <= oHBS.HybridShapes.Count; j++)  //得到转换点
                {
                    hspcEle = (HybridShapePointCoord)oHBS.HybridShapes.Item(j);
                    dPoint[0] = hspcEle.X.Value - dPointOrginEle[0] + dPointOrgin[0];
                    dPoint[1] = hspcEle.Y.Value - dPointOrginEle[1] + dPointOrgin[1];

                    HybridShapePointCoord hspT = oHSFOut.AddNewPointCoord(dPoint[0], dPoint[1], 0);
                    hspT.set_Name(hspcEle.get_Name());
                    oHBSurf3.AppendHybridShape(hspT);


                    alPoint.Add(hspT);
                }

                alLine.Clear();
                for (j = 0; j < alPoint.Count; j++)  //得到连线
                {
                    Reference rf1=oPartOut.CreateReferenceFromObject((HybridShapePointCoord)alPoint[j]);
                    Reference rf2;
                    if (j == alPoint.Count-1)
                    {
                        rf2=oPartOut.CreateReferenceFromObject((HybridShapePointCoord)alPoint[0]);
                    }
                    else
                    {
                        rf2=oPartOut.CreateReferenceFromObject((HybridShapePointCoord)alPoint[j+1]);
                    }

                    HybridShapeLinePtPt hslppT = oHSFOut.AddNewLinePtPt(rf1,rf2);
                    hslppT.set_Name("L"+j.ToString());
                    oHBSurf3.AppendHybridShape(hslppT);


                    alLine.Add(hslppT);
                }

                //形成面
                if (formXYOption.checkBoxSCM.Checked)
                {
                    HybridShapeFill oFill=oHSFOut.AddNewFill();
                    for (j = 0; j < alLine.Count;j++ )
                        oFill.AddBound((Reference)alLine[j]);

                    oFill.set_Name("Fill");
                    oHBSurf3.AppendHybridShape(oFill);

                }

                //原点重置
                dPointOrgin[0] += (double)formXYOption.numericUpDownHJJ.Value+dMaxXLength;
                k1++;
                if (dYLength < dMinYLength)
                {
                    dYLength = dMinYLength;
                }

                toolStripProgressBarCATIA.Value = i;

            }
            toolStripProgressBarCATIA.Value = toolStripProgressBarCATIA.Minimum;
            toolStripStatusLabelCATIA.Text = "导出完毕.";

            oPartOut.Update();


            if (formXYOption.checkBoxCT.Checked) //出图
            {
                buttonCT_Click(null, null);
                

                //Drafting(oHB, oHB2, oHB3);
            }



        }

        //出图 旋转，排列
        private void Drafting(MECMOD.HybridBody hbO, MECMOD.HybridBody hbR, MECMOD.HybridBody hbA)
        {
            double dPaperX = 0, dPaperY = 0;
            double[] dPointOrgin = new double[3];
            double[] dPoint = new double[3];
            int i, j, k;

            dPointOrgin[0] = (double)formXYOption.numericUpDownX.Value;
            dPointOrgin[1] = (double)formXYOption.numericUpDownY.Value;

            double[] dPointOrginEle = new double[3];
            double dMaxX = 0, dMaxXLength = 0;
            double dMinY = 0, dMinYLength = 0, dYLength = 0;

            DrawingDocument ddCatia = (DrawingDocument)CATIA.Documents.Add("Drawing");
            ddCatia.Standard = CatDrawingStandard.catISO;
            DrawingSheets dssCatia = ddCatia.Sheets;
            DrawingSheet dsCatia;

            if (dssCatia.Count >= 1)
            {
                for (i = 1; i <= dssCatia.Count; i++)
                {
                    dssCatia.Remove(i);;
                }
            }
            
            //初始化
            dsCatia = AddNewDSheet(dssCatia, formDrawingsOption.textBoxTM.Text + "1");
            cCATIA.fitWindows();

            //得到纸大小
            dPaperX = dsCatia.GetPaperWidth();
            dPaperY = dsCatia.GetPaperHeight();


            /*////////////////////////////////////////////////////////////
            DrawingView dv1 = dsCatia.Views.Add(hbA.HybridBodies.Item(1).get_Name());

            dv1.Scale = dsCatia.Scale;


            DrawingViewGenerativeBehavior oFrontViewGB;
            oFrontViewGB = dv1.GenerativeBehavior;

            oFrontViewGB.Document = hbA.HybridBodies.Item(5);
            oFrontViewGB.DefineFrontView(1, 0, 0, 0, 1, 0);
            dv1.x = dPaperX;
            dv1.y = dPaperY;
            

            //DrawingText dt = dv1.Texts.Add(hbA.HybridBodies.Item(1).get_Name(),-10,-30);
            //dt.set_Text("111111");
            //dt.x = -10; dt.y = -30;
            oFrontViewGB.Update();

            DrawingTexts dts = dv1.Texts;

            MECMOD.HybridBody os = hbA.HybridBodies.Item(5);
            int ii;
            
            for (ii = 1; ii <= os.HybridShapes.Count; ii++)
            {
                HybridShapePointCoord h = (HybridShapePointCoord)os.HybridShapes.Item(ii);

                DrawingText dt2 = dts.Add(h.get_Name(), h.X.Value, h.Y.Value);
                //dt.set_Text("111111");
                //dt2.x = h.X.Value; dt2.y = h.Y.Value;

            }

            oFrontViewGB.Update();



            ////////////////////////////////////////////////////////////////*/


            //图边缘 上，左，右，下
            double dUp = (double)formDrawingsOption.numericUpDownSBY.Value, dLeft = (double)formDrawingsOption.numericUpDownZBY.Value, dRight = (double)formDrawingsOption.numericUpDownYBY.Value, dDown = (double)formDrawingsOption.numericUpDownXBY.Value;
            double dDownR = (double)formDrawingsOption.numericUpDownJX.Value, dRightR = (double)formDrawingsOption.numericUpDownBKD.Value;//表格间隙，表格宽
            double dTextHeight = (double)formDrawingsOption.numericUpDownBTZG.Value; //字高
            double dYPos = 0;//Y坐标
            double dHJJ = (double)formDrawingsOption.numericUpDownHJJ.Value, dLJJ = (double)formDrawingsOption.numericUpDownLJJ.Value;

            double dWidth = 0;//可写图宽
            dWidth = dPaperX - dRightR - dDownR - dRight - dLeft;
            double dLength = 0;//可写图高
            dLength = dPaperY - dUp - dDown;

            //绘图

            int iStartP = 1, iEndP = 1, iRows=0, iStarP=1; //表格计数
            MECMOD.HybridBody oHBSuf;
            MECMOD.HybridBody oHBSuf1;

            int iNumBodys = 1;
            dYPos = dPaperY - dUp; //Y坐标
            bool bFirstLine = true; //第一行必写
            bool bWriteTable = false; //是否写表
            bool bNewPaper = false; //是否打开新图
            int iNo = 2;

            DrawingView dvTable;//表格视图
            DrawingTable dtTable;
            double cWidth = (double)formDrawingsOption.numericUpDownDYK.Value, rHeight = (double)formDrawingsOption.numericUpDownDYG.Value;
            DrawingViewGenerativeBehavior oFrontViewGB;

            
            while (true)
            {
                if (iNumBodys > hbR.HybridBodies.Count)
                    break;

                //if (iNumBodys == 17)
                //{
                //    int ii = 1;
                //}

                oHBSuf=hbR.HybridBodies.Item(iNumBodys);
                if (oHBSuf.HybridShapes.Count<=2)
                {
                    iNumBodys++;
                    continue;
                }




                //得到一行
                iStartP = iNumBodys; iEndP = iStartP;
                //第一个平面
                drawingCATIA.Rows.Clear();//坐标table

                HybridShapePointCoord hspcEle = (HybridShapePointCoord)oHBSuf.HybridShapes.Item(1);

                object[] oTemp2 = new object[7];
                oTemp2[0] = iStartP;
                oTemp2[1] = (hspcEle.X.Value * dsCatia.Scale);
                oTemp2[2] = (hspcEle.Y.Value * dsCatia.Scale);
                oTemp2[3] = (hspcEle.X.Value * dsCatia.Scale);
                oTemp2[4] = (hspcEle.Y.Value * dsCatia.Scale);
                oTemp2[5] = 0;
                oTemp2[6] = 0;
                drawingCATIA.Rows.Add(oTemp2);

                for (j = 2; j <= oHBSuf.HybridShapes.Count; j++)
                {
                    hspcEle = (HybridShapePointCoord)oHBSuf.HybridShapes.Item(j);
                    if (hspcEle.X.Value * dsCatia.Scale < double.Parse(drawingCATIA.Rows[drawingCATIA.Rows.Count - 1][1].ToString()))
                    {
                        drawingCATIA.Rows[drawingCATIA.Rows.Count - 1][1] = hspcEle.X.Value * dsCatia.Scale;
                    }
                    if (hspcEle.Y.Value * dsCatia.Scale < double.Parse(drawingCATIA.Rows[drawingCATIA.Rows.Count - 1][2].ToString()))
                    {
                        drawingCATIA.Rows[drawingCATIA.Rows.Count - 1][2] = hspcEle.Y.Value * dsCatia.Scale;
                    }
                    if (hspcEle.X.Value * dsCatia.Scale > double.Parse(drawingCATIA.Rows[drawingCATIA.Rows.Count - 1][3].ToString()))
                    {
                        drawingCATIA.Rows[drawingCATIA.Rows.Count - 1][3] = hspcEle.X.Value * dsCatia.Scale;
                    }
                    if (hspcEle.Y.Value * dsCatia.Scale > double.Parse(drawingCATIA.Rows[drawingCATIA.Rows.Count - 1][4].ToString()))
                    {
                        drawingCATIA.Rows[drawingCATIA.Rows.Count - 1][4] = hspcEle.Y.Value * dsCatia.Scale;
                    }

                }

                dMaxX = double.Parse(drawingCATIA.Rows[drawingCATIA.Rows.Count - 1][3].ToString()) - double.Parse(drawingCATIA.Rows[drawingCATIA.Rows.Count - 1][1].ToString());
                dMinY = double.Parse(drawingCATIA.Rows[drawingCATIA.Rows.Count - 1][4].ToString()) - double.Parse(drawingCATIA.Rows[drawingCATIA.Rows.Count - 1][2].ToString());

                for (i = iStartP + 1; i <= hbR.HybridBodies.Count; i++)
                {
                    oHBSuf = hbR.HybridBodies.Item(i);
                    object[] oTemp3 = new object[7];
                    HybridShapePointCoord hspcEle1 = (HybridShapePointCoord)oHBSuf.HybridShapes.Item(1);

                    oTemp3[0] = i;
                    oTemp3[1] = hspcEle1.X.Value * dsCatia.Scale;
                    oTemp3[2] = hspcEle1.Y.Value * dsCatia.Scale;
                    oTemp3[3] = hspcEle1.X.Value * dsCatia.Scale;
                    oTemp3[4] = hspcEle1.Y.Value * dsCatia.Scale;
                    oTemp3[5] = 0;
                    oTemp3[6] = 0;

                    //得到范围
                    for (j = 2; j <= oHBSuf.HybridShapes.Count; j++)
                    {
                        hspcEle1 = (HybridShapePointCoord)oHBSuf.HybridShapes.Item(j);
                        if (hspcEle1.X.Value * dsCatia.Scale < (double)oTemp3[1])
                        {
                            oTemp3[1] = hspcEle1.X.Value * dsCatia.Scale;
                        }
                        if (hspcEle1.Y.Value * dsCatia.Scale < (double)oTemp3[2])
                        {
                            oTemp3[2] = hspcEle1.Y.Value * dsCatia.Scale;
                        }
                        if (hspcEle1.X.Value * dsCatia.Scale > (double)oTemp3[3])
                        {
                            oTemp3[3] = hspcEle1.X.Value * dsCatia.Scale;
                        }
                        if (hspcEle1.Y.Value * dsCatia.Scale > (double)oTemp3[4])
                        {
                            oTemp3[4] = hspcEle1.Y.Value * dsCatia.Scale;
                        }


                    }

                    if (dMaxX + (double)oTemp3[3] - (double)oTemp3[1] + dLJJ >= dWidth) //换行
                    {
                        break;
                    }
                    else //不换行，继续
                    {
                        dMaxX += (double)oTemp3[3] - (double)oTemp3[1] + dLJJ;
                        if (dMinY < (double)oTemp3[4] - (double)oTemp3[2])
                        {
                            dMinY = (double)oTemp3[4] - (double)oTemp3[2];
                        }
                        drawingCATIA.Rows.Add(oTemp3);
                    }

                }

                if (!bFirstLine) //不是第一行，看是否出界
                {
                    if (dYPos - dMinY - dHJJ - dTextHeight < dDown) //出界（非第一行），这行不写
                    {
                        bNewPaper = true;
                        bWriteTable = true;

                        iEndP = iStartP - 1;

                    }
                    else //没出界
                    {
                        //dYPos -= dMinY + dHJJ + dTextHeight;
                        iEndP = i - 1;

                    }
                }
                else //第一行，建立表格视图
                {
                    iStarP = iStartP;
                    bFirstLine = false;
                    bWriteTable = false;
                    
                    iEndP = i - 1;
                }

                //if (i - 1 - iStartP < 1) //第一个面溢出X边界，Y边界
                //    iEndP = iStartP;
                //else

                //iEndP = i - 1;
                
                iNumBodys = iEndP + 1;

                if (iEndP >= hbR.HybridBodies.Count) //最后一个面，写表
                    bWriteTable = true;

                //计算每个面坐标原点
                if (!bNewPaper)
                {
                    dMaxX = dLeft;
                    for (j = 0; j < drawingCATIA.Rows.Count; j++)
                    {
                        drawingCATIA.Rows[j][5] = dMaxX + (double.Parse(drawingCATIA.Rows[j][3].ToString()) - double.Parse(drawingCATIA.Rows[j][1].ToString())) / 2.0;
                        dMaxX += (double.Parse(drawingCATIA.Rows[j][3].ToString()) - double.Parse(drawingCATIA.Rows[j][1].ToString())) + dLJJ;

                        drawingCATIA.Rows[j][6] = dYPos - dMinY - dTextHeight + (double.Parse(drawingCATIA.Rows[j][4].ToString()) - double.Parse(drawingCATIA.Rows[j][2].ToString())) / 2.0;
                    }

                    for (j = 0; j < drawingCATIA.Rows.Count; j++)
                    {

                        //画面
                        DrawingView dvSurf = dsCatia.Views.Add(hbA.HybridBodies.Item(int.Parse(drawingCATIA.Rows[j][0].ToString())).get_Name());
                        dvSurf.Scale = dsCatia.Scale;
                        oFrontViewGB = dvSurf.GenerativeBehavior;

                        oFrontViewGB.Document = hbA.HybridBodies.Item(int.Parse(drawingCATIA.Rows[j][0].ToString()));
                        oFrontViewGB.PointsProjectionMode = CatPointsProjectionMode.catPointsProjectionModeOn;
                        oFrontViewGB.PointsSymbol = (short)(formDrawingsOption.comboBoxFH.SelectedIndex+1);

                        oFrontViewGB.DefineFrontView(1, 0, 0, 0, 1, 0);
                        dvSurf.x = double.Parse(drawingCATIA.Rows[j][5].ToString());
                        dvSurf.y = double.Parse(drawingCATIA.Rows[j][6].ToString());
                        oFrontViewGB.Update();


                        //写点
                        oHBSuf = hbR.HybridBodies.Item(int.Parse(drawingCATIA.Rows[j][0].ToString()));
                        oHBSuf1 = hbA.HybridBodies.Item(int.Parse(drawingCATIA.Rows[j][0].ToString()));

                        //面名称
                        DrawingTexts dtsSurf = dvSurf.Texts;
                        //改写
                        if (dtsSurf.Count == 1)
                        {
                            DrawingText dt1 = dtsSurf.Item(1);
                            dt1.set_Text(oHBSuf1.get_Name());

                            cCATIA.changeFont(formDrawingsOption.fontCatia, ref dt1);


                        }
                        //DrawingText dtSurf = dtsSurf.Add(hbA.HybridBodies.Item(int.Parse(drawingCATIA.Rows[j][0].ToString())).get_Name(), 0, 0);
                        //dtSurf.x = 0; dtSurf.y = 0;




                        for (k = 1; k <= oHBSuf.HybridShapes.Count; k++)
                        {
                            HybridShapePointCoord h = (HybridShapePointCoord)oHBSuf1.HybridShapes.Item(k);
                            DrawingText dt2 = dtsSurf.Add(h.get_Name(), h.X.Value, h.Y.Value);

                            cCATIA.changeFont(formDrawingsOption.fontCatia1, ref dt2);


                        }


                        //oFrontViewGB.Update();



                    }
                }

                //写表
                if (bWriteTable)
                {
                    iRows = 2;//表头
                    //计算表行数
                    for (j = iStarP; j <= iEndP; j++)
                    {
                        iRows++;
                        oHBSuf = hbR.HybridBodies.Item(j);
                        iRows += oHBSuf.HybridShapes.Count;
                    }
                    

                    //表头
                    DrawingText dtC;
                    dvTable = dsCatia.Views.Add("Table");
                    dtTable = dvTable.Tables.Add(dPaperX - dRightR - dRight, dPaperY - dUp, iRows, 6, rHeight, cWidth);
                    dtTable.MergeCells(1, 1, 1, 6);
                    dtTable.SetCellAlignment(1, 1, CatTablePosition.CatTableMiddleCenter);
                    dtTable.SetCellString(1, 1, formDrawingsOption.textBoxBM.Text);
                    dtC = dtTable.GetCellObject(1, 1);
                    cCATIA.changeFont(formDrawingsOption.fontCatia2, ref dtC);


                    dtTable.SetCellString(2, 1, "N");
                    dtTable.SetCellAlignment(2, 1, CatTablePosition.CatTableMiddleCenter);
                    dtC = dtTable.GetCellObject(1, 1);
                    cCATIA.changeFont(formDrawingsOption.fontCatia2, ref dtC);

                    
                    dtTable.SetCellString(2, 2, "X");
                    dtTable.SetCellAlignment(2, 2, CatTablePosition.CatTableMiddleCenter);
                    dtC = dtTable.GetCellObject(2, 2);
                    cCATIA.changeFont(formDrawingsOption.fontCatia2, ref dtC);

                    
                    dtTable.SetCellString(2, 3, "Y");
                    dtTable.SetCellAlignment(2, 3, CatTablePosition.CatTableMiddleCenter);
                    dtC = dtTable.GetCellObject(2, 3);
                    cCATIA.changeFont(formDrawingsOption.fontCatia2, ref dtC);


                    dtTable.SetCellString(2, 4, "Z");
                    dtTable.SetCellAlignment(2, 4, CatTablePosition.CatTableMiddleCenter);
                    dtC = dtTable.GetCellObject(2, 4);
                    cCATIA.changeFont(formDrawingsOption.fontCatia2, ref dtC);


                    dtTable.SetCellString(2, 5, "X'");
                    dtTable.SetCellAlignment(2, 5, CatTablePosition.CatTableMiddleCenter);
                    dtC = dtTable.GetCellObject(2, 5);
                    cCATIA.changeFont(formDrawingsOption.fontCatia2, ref dtC);


                    dtTable.SetCellString(2, 6, "Y'");
                    dtTable.SetCellAlignment(2, 6, CatTablePosition.CatTableMiddleCenter);
                    dtC = dtTable.GetCellObject(2, 6);
                    cCATIA.changeFont(formDrawingsOption.fontCatia2, ref dtC);


                    iRows = 3;
                    for (j = iStarP; j <= iEndP; j++)
                    {
                        oHBSuf = hbO.HybridBodies.Item(j);
                        oHBSuf1 = hbR.HybridBodies.Item(j);

                        dtTable.MergeCells(iRows, 1, 1, 6);
                        dtTable.SetCellString(iRows, 1, oHBSuf.get_Name());
                        dtTable.SetCellAlignment(iRows, 1, CatTablePosition.CatTableMiddleCenter);
                        dtC = dtTable.GetCellObject(iRows, 1);
                        cCATIA.changeFont(formDrawingsOption.fontCatia2, ref dtC);
                        iRows++;

                        for (k = 1; k <= oHBSuf.HybridShapes.Count; k++)
                        {
                            HybridShapePointCoord h2 = (HybridShapePointCoord)oHBSuf.HybridShapes.Item(k);
                            HybridShapePointCoord h3 = (HybridShapePointCoord)oHBSuf1.HybridShapes.Item(k);

                            dtTable.SetCellString(iRows, 1, h2.get_Name());
                            dtTable.SetCellAlignment(iRows, 1, CatTablePosition.CatTableMiddleCenter);
                            dtC = dtTable.GetCellObject(iRows, 1);
                            cCATIA.changeFont(formDrawingsOption.fontCatia2,ref dtC);

                            dtTable.SetCellString(iRows, 2, h2.X.Value.ToString("f"+formDrawingsOption.numericUpDownJD.Value.ToString()));
                            dtTable.SetCellAlignment(iRows, 2, CatTablePosition.CatTableMiddleCenter);
                            dtC = dtTable.GetCellObject(iRows, 2);
                            cCATIA.changeFont(formDrawingsOption.fontCatia2, ref dtC);

                            dtTable.SetCellString(iRows, 3, h2.Y.Value.ToString("f" + formDrawingsOption.numericUpDownJD.Value.ToString()));
                            dtTable.SetCellAlignment(iRows, 3, CatTablePosition.CatTableMiddleCenter);
                            dtC = dtTable.GetCellObject(iRows, 3);
                            cCATIA.changeFont(formDrawingsOption.fontCatia2, ref dtC);

                            dtTable.SetCellString(iRows, 4, h2.Z.Value.ToString("f" + formDrawingsOption.numericUpDownJD.Value.ToString()));
                            dtTable.SetCellAlignment(iRows, 4, CatTablePosition.CatTableMiddleCenter);
                            dtC = dtTable.GetCellObject(iRows, 4);
                            cCATIA.changeFont(formDrawingsOption.fontCatia2, ref dtC);

                            dtTable.SetCellString(iRows, 5, h3.X.Value.ToString("f" + formDrawingsOption.numericUpDownJD.Value.ToString()));
                            dtTable.SetCellAlignment(iRows, 5, CatTablePosition.CatTableMiddleCenter);
                            dtC = dtTable.GetCellObject(iRows, 5);
                            cCATIA.changeFont(formDrawingsOption.fontCatia2, ref dtC);

                            dtTable.SetCellString(iRows, 6, h3.Y.Value.ToString("f" + formDrawingsOption.numericUpDownJD.Value.ToString()));
                            dtTable.SetCellAlignment(iRows, 6, CatTablePosition.CatTableMiddleCenter);
                            dtC = dtTable.GetCellObject(iRows, 6);
                            cCATIA.changeFont(formDrawingsOption.fontCatia2, ref dtC);


                            iRows++;

                        }

                    }

                }
                dYPos -= dMinY + dHJJ + dTextHeight;
                if (bNewPaper) //
                {
                    //加新图
                    dsCatia = AddNewDSheet(dssCatia, formDrawingsOption.textBoxTM.Text + iNo.ToString());
                    iNo++;

                    cCATIA.fitWindows();

                    dYPos = dPaperY - dUp; //Y坐标
                    bFirstLine = true; //第一行必写
                    bNewPaper = false;
                }


                
                
            }
            //oFrontViewGB.Update();

            



        }

        private DrawingSheet AddNewDSheet(DrawingSheets dssCatia, string sName)
        {
            DrawingSheet dsCatia = dssCatia.Add(sName);

            //纸张大小
            switch (formDrawingsOption.comboBoxZZ.Text)
            {
                case "A0":
                    dsCatia.PaperSize = DRAFTINGITF.CatPaperSize.catPaperA0;
                    break;
                case "A1":
                    dsCatia.PaperSize = DRAFTINGITF.CatPaperSize.catPaperA1;
                    break;
                case "A2":
                    dsCatia.PaperSize = DRAFTINGITF.CatPaperSize.catPaperA2;
                    break;
                case "A3":
                    dsCatia.PaperSize = DRAFTINGITF.CatPaperSize.catPaperA3;
                    break;
                case "A4":
                    dsCatia.PaperSize = DRAFTINGITF.CatPaperSize.catPaperA4;
                    break;
                default:
                    dsCatia.PaperSize = DRAFTINGITF.CatPaperSize.catPaperA3;
                    break;
            }

            //比例
            dsCatia.Scale = 1.0 / (double)(formDrawingsOption.numericUpDownBL.Value);

            //方向
            if (formDrawingsOption.radioButtonHX.Checked)
                dsCatia.Orientation = DRAFTINGITF.CatPaperOrientation.catPaperLandscape;
            else
                dsCatia.Orientation = DRAFTINGITF.CatPaperOrientation.catPaperPortrait;

            return dsCatia;

        }



        private double DotProduct(object[] vect1, object[] vect2)
        {
            return ((double)vect1[0] * (double)vect2[0] + (double)vect1[1] * (double)vect2[1] + (double)vect1[2] * (double)vect2[2]);
        }

        private bool NormalizeVector(object[] invect, ref object[] normvect)
        {
            double mag;
            mag = Math.Sqrt((double)invect[0] * (double)invect[0] + (double)invect[1] * (double)invect[1] + (double)invect[2] * (double)invect[2]);

            if(mag<=EP)
                return false;

            normvect[0] = (double)invect[0] / mag;
            normvect[1] = (double)invect[1] / mag;
            normvect[2] = (double)invect[2] / mag;

            return true;

        }

        HybridShapePointCoord getXYPlaneCoord(HybridShapeFactory hsf, HybridShapePointCoord P1, HybridShapePointCoord P2)
        {
            double[] dPtemp = new double[3];

            //得到相对坐标
            dPtemp[0] = P2.X.Value - P1.X.Value;
            dPtemp[1] = P2.Y.Value - P1.Y.Value;
            dPtemp[2] = P2.Z.Value - P1.Z.Value;

            HybridShapePointCoord PointXY;
            if (Math.Abs(dPtemp[0]) < EP && Math.Abs(dPtemp[1]) < EP && Math.Abs(dPtemp[2]) < EP)
            {
                 PointXY = oHSFOut.AddNewPointCoord(dPtemp[0], dPtemp[1], 0);
                 return PointXY;
            }

            double dAngle = Math.Atan2(dPtemp[1], dPtemp[0]);
            double dLength = Math.Sqrt(dPtemp[0] * dPtemp[0] + dPtemp[1] * dPtemp[1] + dPtemp[2] * dPtemp[2]);

            PointXY = oHSFOut.AddNewPointCoord(P1.X.Value + dLength * Math.Sin(dAngle), P1.Y.Value+dLength * Math.Cos(dAngle), 0);

            return PointXY;
        }

        private void buttonDGM_Click(object sender, EventArgs e)
        {
            Selection selS = cCATIA.oPartDoc.Selection.Selection;
            cSelection = cCATIA.oPartDoc.Selection;
            int i, j, k;
            int num;

            this.TopMost = false;

            k = 0;
            if (cSelection.Count < 1)
            {
                if (MessageBox.Show("是否对所有点元素更名？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }

                for (i = 1; i <= cCATIA.oPartDoc.Part.HybridBodies.Count; i++)
                {
                    HybridBody myBody = cCATIA.oPart.HybridBodies.Item(i);

                    if (myBody.get_Name().IndexOf(textBoxZMC.Text.Trim()) != -1) //???????????????
                    {
                        num = 0;
                        for (j = 1; j <= myBody.HybridShapes.Count; j++)
                        {
                            HybridShape myShape = myBody.HybridShapes.Item(j);
                            if (myShape.get_Name().IndexOf(textBoxMMC.Text.Trim()) != -1)
                            {
                                selS.Clear();

                                selS.Add((AnyObject)myShape);
                                selS.Search("((((((CATStFreeStyleSearch.Point + CAT2DLSearch.2DPoint) + CATSketchSearch.2DPoint) + CATDrwSearch.2DPoint) + CATPrtSearch.Point) + CATGmoSearch.Point) + CATSpdSearch.Point),sel");

                                for (k = 0; k < selS.Count; k++)
                                {
                                    AnyObject hyPoint = (AnyObject)(selS.Item(k + 1).Value);
                                    if (hyPoint.get_Name().IndexOf(textBoxDMC.Text.Trim()) == -1)
                                        continue;

                                    if(checkBoxDD.Checked) //单独编号
                                        hyPoint.set_Name(getStringNum(k));
                                    else
                                        hyPoint.set_Name(getStringNum(num));

                                    num++;


                                }
 
                            }

                        }
                    }

                }


            }
            else
            {
                for (i = 1; i <= cSelection.Count; i++)
                {
                    HybridBody myBody1;
                    try
                    {
                        myBody1 = (HybridBody)cSelection.Item(i).Value;
                    }
                    catch
                    {
                        continue;
                    }
                    if (myBody1.get_Name().IndexOf(textBoxZMC.Text.Trim()) != -1) //???????????????
                    {
                        num = 0;
                        for (j = 1; j <= myBody1.HybridShapes.Count; j++)
                        {
                            HybridShape myShape1 = myBody1.HybridShapes.Item(j);
                            if (myShape1.get_Name().IndexOf(textBoxMMC.Text.Trim()) != -1)
                            {
                                selS.Clear();

                                selS.Add((AnyObject)myShape1);
                                selS.Search("((((((CATStFreeStyleSearch.Point + CAT2DLSearch.2DPoint) + CATSketchSearch.2DPoint) + CATDrwSearch.2DPoint) + CATPrtSearch.Point) + CATGmoSearch.Point) + CATSpdSearch.Point),sel");

                                for (k = 0; k < selS.Count; k++)
                                {
                                    AnyObject hyPoint = (AnyObject)(selS.Item(k + 1).Value);
                                    if (hyPoint.get_Name().IndexOf(textBoxDMC.Text.Trim()) == -1)
                                        continue;

                                    if (checkBoxDD.Checked) //单独编号
                                        hyPoint.set_Name(getStringNum(k));
                                    else
                                        hyPoint.set_Name(getStringNum(num));

                                    num++;


                                }

                            }

                        }
                    }

                }
            }
            cCATIA.oPart.Update();
            MessageBox.Show("命名更新完毕");
        }

        private void buttonZNJG_Click(object sender, EventArgs e)
        {
            if (cElement.Count < 1)
            {
                MessageBox.Show("请先选定统计元素", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }



            int i, j, k, no;
            Workbench TheSPAWorkbench;
            TheSPAWorkbench = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench");

            Measurable TheMeasurable;

            dtCATIA.Clear();
            pointElement.Clear();
            sufElement.Clear();

            //Selection selS = cCATIA.oPartDoc.Selection;
            Selection selS = cCATIA.oPartDoc.Selection.Selection;
            bool bAdd = false;
            no = 1;


            for (i = 0; i < cElement.Count; i++)
            {
                //try
                //{

                HybridBody hybHB;

                hybHB = (HybridBody)cElement[i];

                toolStripStatusLabelCATIA.Text = "正在导出顶点....";
                toolStripProgressBarCATIA.Maximum = hybHB.HybridShapes.Count;


                for (j = 1; j <= hybHB.HybridShapes.Count; j++)
                {
                    try
                    {
                        HybridShape hbsSHAPE = (HybridShape)hybHB.HybridShapes.Item(j);


                        if (hbsSHAPE.get_Name().IndexOf(textBoxMMC.Text.Trim()) != -1)
                        {
                            Reference refSuf = cCATIA.oPart.CreateReferenceFromObject(hbsSHAPE);
                            //HybridShapeBoundary hsbSufB = cCATIA.oHSF.AddNewBoundaryOfSurface(refSuf);
                            //hbsSHAPE.AppendHybridShape(hsbSufB);

                            selS.Clear();
                            selS.Add((AnyObject)hbsSHAPE);
                            selS.Search("Topology.CGMVertex,sel");
                            bAdd = false;

                            for (k = 0; k < selS.Count; k++)
                            {
                                AnyObject hyPoint = (AnyObject)(selS.Item(k + 1).Value);

                                if (!bAdd)
                                {
                                    sufElement.Add(hbsSHAPE);
                                    bAdd = true;
                                }
                                pointElement.Add(hyPoint);

                                //Reference refP = cCATIA.oPart.CreateReferenceFromObject(hyPoint);

                                TheMeasurable = ((SPAWorkbench)TheSPAWorkbench).GetMeasurable((Reference)hyPoint);



                                object[] oPoint = new object[3];

                                TheMeasurable.GetPoint(oPoint);

                                //hyP.GetCoordinates((Array)oPoint);

                                object[] oTemp = new object[8];

                                oTemp[0] = no.ToString();
                                oTemp[1] = hbsSHAPE.get_Name();
                                if(checkBoxDD.Checked) //单独编号
                                    oTemp[2]=getStringNum(k);
                                else
                                    oTemp[2]=getStringNum(no-1);

                                oTemp[3] = oPoint[0].ToString();
                                oTemp[4] = oPoint[1].ToString();
                                oTemp[5] = oPoint[2].ToString();
                                oTemp[6] = sufElement.Count - 1;
                                oTemp[7] = pointElement.Count - 1;

                                if (checkBoxXRDD.Checked)//写入顶点
                                {
                                    HybridShapePointCoord hspcT = cCATIA.oHSF.AddNewPointCoord((double)oPoint[0],(double)oPoint[1],(double)oPoint[2]);
                                    hspcT.set_Name(oTemp[2].ToString());
                                    hbsSHAPE.AppendHybridShape(hspcT);
                                }


                                no++;



                                dtCATIA.Rows.Add(oTemp);
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                    toolStripProgressBarCATIA.Value = j;
                }

                //}
                //catch
                //{
                //    continue;
                //}

            }
            cCATIA.oPart.Update();
            MessageBox.Show("顶点导出完毕，请选择打印或输出到EXCEL", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            toolStripStatusLabelCATIA.Text = "顶点导出完毕";
            toolStripProgressBarCATIA.Value = toolStripProgressBarCATIA.Minimum;
        }

        private void buttonCT_Click(object sender, EventArgs e)
        {
            formDrawingsOption.ShowDialog();
            if (formDrawingsOption.bCancel)
                return;
            try
            {
                MECMOD.PartDocument oPartD = (MECMOD.PartDocument)CATIA.ActiveDocument;
                Drafting(oPartD.Part.HybridBodies.Item(1), oPartD.Part.HybridBodies.Item(2), oPartD.Part.HybridBodies.Item(3));
            }
            catch(Exception ex)
            {
                MessageBox.Show("读取错误，请重新打开转换文件");
                return;
            }
            MessageBox.Show("出图完毕，可以手动调整");
        }
    }
}
