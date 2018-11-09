using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using INFITF;
using MECMOD;
using HybridShapeTypeLib;
using PARTITF;
using ProductStructureTypeLib;
using DRAFTINGITF;
using System.IO;

namespace ColumnToExcel
{
    class ClassCATIA
    {
        public INFITF.Application CATIA;
        public INFITF.Documents docCATIA;
        public ProductDocument oProductDoc;
        public MECMOD.PartDocument oPartDoc;
        public DrawingDocument oDrawingDoc;
        public MECMOD.Part oPart;

        public MECMOD.Bodies oBodies;
        public MECMOD.Body oBody;

        public MECMOD.HybridBodies oHBodies;
        public MECMOD.HybridBody oHBody;

        public ShapeFactory oSF;
        public HybridShapeFactory oHSF;
        
        
        MECMOD.Sketches skCATIA;

        private const double EP = 1e-3;

        public bool InitCATIAPart()
        {
            try
            {
                oPartDoc = (MECMOD.PartDocument)CATIA.ActiveDocument;
                oPart = oPartDoc.Part;
                oBodies = oPart.Bodies;
                oBody = oPart.MainBody;
                oHBodies = oPart.HybridBodies;

                oSF = (ShapeFactory)oPart.ShapeFactory;
                oHSF = (HybridShapeFactory)oPart.HybridShapeFactory;
            }
            catch
            {
            }
            return true;
        }

        public bool InitCATIAPart(bool bNewPart, string strPart)
        {
            if (bNewPart) 
            {
                //初始化
                docCATIA = CATIA.Documents;
                oPartDoc = (MECMOD.PartDocument)docCATIA.Add("Part");
            }
            else
            {
                if(strPart.Trim()=="")
                {
                    oPartDoc=(MECMOD.PartDocument)CATIA.ActiveDocument;
                    if(oPartDoc==null)
                    {
                        docCATIA = CATIA.Documents;
                        oPartDoc = (MECMOD.PartDocument)docCATIA.Add("Part");
                    }
                }
                else
                {
                    if(System.IO.File.Exists(strPart)) //有文件
                    {
                         oPartDoc =(MECMOD.PartDocument)CATIA.Documents.Open(strPart);
                    }
                    else
                    {
                        return false;
                    }
                
                }
            }

            oPart=oPartDoc.Part;
            oBodies=oPart.Bodies;
            oBody=oPart.MainBody;
            oHBodies=oPart.HybridBodies;

            oSF=(ShapeFactory)oPart.ShapeFactory;
            oHSF=(HybridShapeFactory)oPart.HybridShapeFactory;

            return true;
            
        }

        public bool InitCATIADrawing(bool bNewDrawing, string strDarwins)
        {
            if (bNewDrawing)
            {
                //初始化
                docCATIA = CATIA.Documents;
                oDrawingDoc = (DrawingDocument)docCATIA.Add("Drawing");
            }
            else
            {
                if (strDarwins.Trim() == "")
                {
                    oDrawingDoc = (DrawingDocument)CATIA.ActiveDocument;
                    if (oDrawingDoc == null)
                    {
                        docCATIA = CATIA.Documents;
                        oDrawingDoc = (DrawingDocument)docCATIA.Add("Drawing");
                    }
                }
                else
                {
                    if (System.IO.File.Exists(strDarwins)) //有文件
                    {
                        oDrawingDoc = (DrawingDocument)CATIA.Documents.Open(strDarwins);
                    }
                    else
                    {
                        return false;
                    }

                }
            }

            return true;
            
        }

        public MECMOD.HybridBody AddHBody(string HBodyName)
        {
            MECMOD.HybridBody oHB;            

            try
            {
                oHB = oHBodies.Add();
                if (HBodyName != "")
                {
                    oHB.set_Name(HBodyName);
                }

                return oHB;
            }
            catch
            {
                return null;
            }
        }

        public void HideShow(object Element, int isShow)
        {
            Reference refElement;

            refElement = oPart.CreateReferenceFromObject((INFITF.AnyObject)Element);

            oHSF.GSMVisibility(refElement, isShow);
        }

        public bool InitCATIAProduct(bool bNewProduct, string strProduct)
        {
            if (bNewProduct)
            {
                //初始化
                docCATIA = CATIA.Documents;
                oProductDoc = (ProductDocument)docCATIA.Add("Product");
            }
            else
            {
                if (strProduct.Trim() == "")
                {
                    oProductDoc = (ProductDocument)CATIA.ActiveDocument;
                    if (oProductDoc == null)
                    {
                        docCATIA = CATIA.Documents;
                        oProductDoc = (ProductDocument)docCATIA.Add("Product");
                    }
                }
                else
                {
                    if (System.IO.File.Exists(strProduct)) //有文件
                    {
                        oProductDoc = (ProductDocument)CATIA.Documents.Open(strProduct);
                    }
                    else
                    {
                        return false;
                    }

                }
            }

            return true;
        }

        public bool IsSameNum(double dNum1, double dNum2)
        {
            if (Math.Abs(dNum1 - dNum2) < EP)
                return true;
            else
                return false;
        }

        //将windows字体转化为Catia字体
        public bool changeFont(Font fontWindows,ref DrawingText dtCatia)
        {
            dtCatia.SetFontName(0, 0, fontWindows.Name);
            dtCatia.SetFontSize(0, 0, (double)fontWindows.Size);
            if (fontWindows.Bold)
                dtCatia.SetParameterOnSubString(CatTextProperty.catBold, 0, 0, 1);
            if (fontWindows.Italic)
                dtCatia.SetParameterOnSubString(CatTextProperty.catItalic, 0, 0, 1);
            if (fontWindows.Underline)
                dtCatia.SetParameterOnSubString(CatTextProperty.catUnderline, 0, 0, 1);
            return true;
        }


        //调整窗口到适合
        public bool fitWindows()
        {
            SpecsAndGeomWindow sAGWCatia;
            sAGWCatia = (SpecsAndGeomWindow)CATIA.ActiveWindow;

            SpecsViewer sViewerCatia;
            sViewerCatia = (SpecsViewer)sAGWCatia.ActiveViewer;

            sViewerCatia.Reframe();

            return true;
        }
    }
}
