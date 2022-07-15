#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;
using System.IO;

#endregion

namespace RevitAddinAcademy_RT
{
    [Transaction(TransactionMode.Manual)]
    public class CmdProjectSetup : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            //Get Excel file
            Forms.OpenFileDialog openFileDialog = new Forms.OpenFileDialog();
            openFileDialog.InitialDirectory = "C:\\";
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Excel files (*.xls)|*xls* | All files (*.*)|*.*";

            string filePath = "";

            if (openFileDialog.ShowDialog() == Forms.DialogResult.OK)
            {
               filePath = openFileDialog.FileName;
            }

            try
            {
                //Setup Excel
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWb = excelApp.Workbooks.Open(filePath);

                //Levels Data
                int index1 = GetExcelWorksheetByName(excelWb, "Levels");
                Excel.Worksheet excelWs1 = excelApp.Worksheets.Item[index1];
                Excel.Range excelRng1 = excelWs1.UsedRange;
                int rowCount1 = excelRng1.Rows.Count;

                List<LevelStruct> levelList = new List<LevelStruct>();

                for (int i = 2; i <= rowCount1; i++)
                {
                    Excel.Range cell1 = excelWs1.Cells[i, 1];
                    Excel.Range cell2 = excelWs1.Cells[i, 2];

                    LevelStruct levelStruct = new LevelStruct();
                    levelStruct.levelName = cell1.Value.ToString();
                    levelStruct.levelElev = cell2.Value;

                    levelList.Add(levelStruct);
                }

                //Sheet Data
                int index2 = GetExcelWorksheetByName(excelWb, "Sheets");
                Excel.Worksheet excelWs2 = excelApp.Worksheets.Item[index2];
                Excel.Range excelRng2 = excelWs2.UsedRange;
                int rowCount2 = excelRng2.Rows.Count;

                List<SheetStruct> sheetList = new List<SheetStruct>();

                for (int j = 2; j <= rowCount1; j++)
                {
                    Excel.Range cell1 = excelWs2.Cells[j, 1];
                    Excel.Range cell2 = excelWs2.Cells[j, 2];
                    Excel.Range cell3 = excelWs2.Cells[j, 3];
                    Excel.Range cell4 = excelWs2.Cells[j, 4];
                    Excel.Range cell5 = excelWs2.Cells[j, 5];

                    SheetStruct sheetStruct = new SheetStruct();
                    sheetStruct.sheetNum = cell1.Value.ToString();
                    sheetStruct.sheetName = cell2.Value.ToString();
                    sheetStruct.viewName = cell3.Value.ToString();
                    sheetStruct.drawnBy = cell4.Value.ToString();
                    sheetStruct.checkedBy = cell5.Value.ToString();

                    sheetList.Add(sheetStruct);
                }

                //Select ViewFamily types for creating Floor Plan and RCP views
                FilteredElementCollector collector1 = new FilteredElementCollector(doc);
                collector1.OfClass(typeof(ViewFamilyType));

                ViewFamilyType curVFT = null;
                ViewFamilyType CurRCPVFT = null;
                foreach(ViewFamilyType curElem in collector1)
                {
                    if(curElem.ViewFamily == ViewFamily.FloorPlan)
                    {
                        curVFT = curElem;
                    }
                    else if(curElem.ViewFamily == ViewFamily.CeilingPlan)
                    {
                        CurRCPVFT = curElem;
                    }
                }


                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Project Setup");

                    //Process levels
                    for (int i = 0; i < levelList.Count; i++)
                    {
                        LevelStruct curLevel = levelList[i];

                        try
                        {
                            Level newLevel = Level.Create(doc, curLevel.levelElev);
                            newLevel.Name = curLevel.levelName;

                            ViewPlan curPlan = ViewPlan.Create(doc, curVFT.Id, newLevel.Id);
                            curPlan.Name = curLevel.levelName;

                            ViewPlan curRCP = ViewPlan.Create(doc, CurRCPVFT.Id, newLevel.Id);
                            curRCP.Name = curLevel.levelName + " RCP";
                        }
                        catch (Exception ex)
                        {
                            Debug.Print(ex.Message);
                        }
                    }

                    //Process Sheets
                    for (int j = 0; j < sheetList.Count; j++)
                    {
                        SheetStruct curSheet = sheetList[j];
                        ElementId titleBlock = GetTitleblockByName(doc, "Alliance 30 x 42");
                        XYZ insPoint = new XYZ(0,0,0);
                        View viewAdd = GetViewByName(doc, curSheet.viewName);

                        try
                        {
                            ViewSheet newSheet = ViewSheet.Create(doc, titleBlock);
                            newSheet.Name = curSheet.sheetName;
                            newSheet.SheetNumber = curSheet.sheetNum;

                            SetParameterValueByName(newSheet, "Drawn By", curSheet.drawnBy);
                            SetParameterValueByName(newSheet, "Checked By", curSheet.checkedBy);

                            if (viewAdd != null)
                            {
                                Viewport newVP = Viewport.Create(doc, newSheet.Id, viewAdd.Id, insPoint);
                            }
                            else
                            {
                                TaskDialog.Show("Error", "View not found.");
                            }


                        }
                        catch (Exception ex)
                        {
                            Debug.Print(ex.Message);
                        }
                    }

                    tx.Commit();
                }

                excelWb.Close();
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
            }


            return Result.Succeeded;
        }

        internal struct LevelStruct
        {
            public string levelName;
            public double levelElev;
        }

        internal struct SheetStruct
        {
            public string sheetNum;
            public string sheetName;
            public string viewName;
            public string drawnBy;
            public string checkedBy;
        }

        internal View GetViewByName(Document doc, string viewName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(View));

            foreach(View curView in collector)
            {
                if(curView.Name == viewName)
                {
                    return curView;
                }
            }

            return null;
        }

        public int GetExcelWorksheetByName(Excel.Workbook excelWb, string name)
        {
            int count = excelWb.Worksheets.Count;
            int index = 1;

            for (int i = 1; i <= count; i++)
            {
                Excel.Worksheet curWs = excelWb.Worksheets[i];
                if(curWs.Name == name)
                {
                    index = curWs.Index;
                }
            }

            return index;
        }
        
        public ElementId GetTitleblockByName(Document doc, string name)
        {
            FilteredElementCollector collector2 = new FilteredElementCollector(doc);
            collector2.OfCategory(BuiltInCategory.OST_TitleBlocks);
            collector2.WhereElementIsElementType();

            string curName = "";
            ElementId elemID = null;

            foreach (Element e in collector2)
            {
                curName = e.Name;
                if(curName == name)
                {
                    elemID = e.Id;
                }
            }
            if(elemID == null)
            {
                elemID = collector2.FirstElementId();
            }

            return elemID;

        }
        
        public void SetParameterValueByName(Element element, string paramName, string paramValue)
        {
            try
            {
                foreach (Parameter curParam in element.Parameters)
                {
                    if (curParam.Definition.Name == paramName)
                    {
                        if(curParam.StorageType is StorageType.String)
                        {
                            curParam.Set(paramValue);
                        }
                        else if(curParam.StorageType is StorageType.Double || curParam.StorageType is StorageType.Integer)
                        {
                            curParam.SetValueString(paramValue);
                        }
                        
                    }
                }
            }
            catch(Exception ex)
            {
                Debug.Print(ex.Message);
            }
            
        }


    }
}
