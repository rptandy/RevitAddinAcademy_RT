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

            string excelFile = @"C:\Users\rebeccat\Documents\_Revit Training\Revit Add-in Academy\Session 02\Session02_Challenge.xlsx";

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);

            Excel.Worksheet excelWs1 = excelApp.Worksheets.Item[1]; //Levels
            Excel.Range excelRng1 = excelWs1.UsedRange;
            int rowCount1 = excelRng1.Rows.Count;
            
            Excel.Worksheet excelWs2 = excelApp.Worksheets.Item[2]; //Sheets
            Excel.Range excelRng2 = excelWs2.UsedRange;
            int rowCount2 = excelRng2.Rows.Count;

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Transaction Name");

                //Process levels
                for(int i = 2; i <= rowCount1; i++)
                {
                    Excel.Range cell1 = excelWs1.Cells[i, 1];
                    Excel.Range cell2 = excelWs1.Cells[i, 2];

                    string levelName = cell1.Value.ToString();
                    double levelElev = cell2.Value;

                    Level newLevel = Level.Create(doc, levelElev);
                    newLevel.Name = levelName;
                }

                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                collector.WhereElementIsElementType();

                //Process Sheets
                for (int j = 2; j <= rowCount2; j++)
                {
                    Excel.Range cell1 = excelWs2.Cells[j, 1];
                    Excel.Range cell2 = excelWs2.Cells[j, 2];

                    string sheetNum = cell1.Value.ToString();
                    string sheetName = cell2.Value.ToString();

                    ViewSheet newSheet = ViewSheet.Create(doc, collector.FirstElementId());
                    newSheet.Name = sheetName;
                    newSheet.SheetNumber = sheetNum;
                }

                tx.Commit();
            }

            excelWb.Close();
            excelApp.Quit();

            return Result.Succeeded;
        }
    }
}
