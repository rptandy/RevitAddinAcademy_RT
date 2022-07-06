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

#endregion

namespace RevitAddinAcademy_RT
{
    [Transaction(TransactionMode.Manual)]
    public class cmdProjectSetup : IExternalCommand
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

            // Get Excel file
            string excelFile = @"C:\Users\rebeccat\Documents\_Revit Training\Revit Add-in Academy\Session 02\Session02_Challenge.xlsx";

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);

            //excel starts numbering at 1, not 0
            //Levels is first sheet [1]
            Excel.Worksheet excelWS1 = excelWb.Worksheets.Item[1];
            Excel.Range excelRng1 = excelWS1.UsedRange;
            int rowCount1 = excelRng1.Rows.Count; 

            //Sheets is second sheet [2]
            Excel.Worksheet excelWS2 = excelWb.Worksheets.Item[2];
            Excel.Range excelRng2 = excelWS2.UsedRange;
            int rowCount2 = excelRng2.Rows.Count; 


            using (Transaction t = new Transaction(doc))
            {
                t.Start("Project setup in Revit");

                for (int i = 2; i <= rowCount1; i++)
                {
                    Excel.Range cell1 = excelWS1.Cells[i, 1];   // Level names
                    Excel.Range cell2 = excelWS1.Cells[i, 2];   // Level elevations

                    string data1 = cell1.Value.ToString();
                    double data2 = Convert.ToDouble(cell2.Value);

                    Level curLevel = Level.Create(doc, data2);
                    curLevel.Name = data1;
                }

                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                collector.WhereElementIsElementType();

                for (int i = 2; i <= rowCount2; i++)
                {
                    Excel.Range cell1 = excelWS2.Cells[i, 1];  // Sheet numbers
                    Excel.Range cell2 = excelWS2.Cells[i, 2];  // Sheet names

                    string data1 = cell1.Value.ToString();
                    string data2 = cell2.Value.ToString();

                    ViewSheet curSheet = ViewSheet.Create(doc, collector.FirstElementId());
                    curSheet.SheetNumber = data1;
                    curSheet.Name = data2;
                }

                t.Commit();
            }
            
            excelWb.Close();
            excelApp.Quit();

            return Result.Succeeded;
        }
    }
}
