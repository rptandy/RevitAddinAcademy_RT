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
    public class Command : IExternalCommand
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

            string excelFile = @"C:\Users\rebeccat\Documents\_Revit Training\Revit Add-in Academy\Session 02\Sheets.xlsx";

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);
            //excel starts numbering at 1, not 0
            Excel.Worksheet excelWs = excelWb.Worksheets.Item[1];

            Excel.Range excelRng = excelWs.UsedRange;
            int rowCount = excelRng.Rows.Count;

            //do some stuff in excel
            List<string[]> dataList = new List<string[]>();
            for (int i = 1; i < rowCount; i++)
            {
                Excel.Range cell1 = excelWs.Cells[i, 1];
                Excel.Range cell2 = excelWs.Cells[i, 2];

                string data1 = cell1.Value.ToString();
                string data2 = cell2.Value.ToString();

                string[] dataArray = new string[2];
                dataArray[0] = data1;
                dataArray[1] = data2;

                dataList.Add(dataArray);


            }

            using(Transaction t = new Transaction(doc))
            {
                t.Start("Create some Revit stuff");
                
                //create level
                Level curLevel = Level.Create(doc, 100);


                 //create sheet

                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                //specify types, not instances
                collector.WhereElementIsElementType();

                ViewSheet curSheet = ViewSheet.Create(doc, collector.FirstElementId());
                curSheet.SheetNumber = "A101";
                curSheet.Name = "New Sheet";

                t.Commit();
            }

            
            excelWb.Close();
            excelApp.Quit();

            return Result.Succeeded;
        }
    }
}
