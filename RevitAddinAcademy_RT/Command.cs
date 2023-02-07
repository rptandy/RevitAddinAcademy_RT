#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;

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

            string modelPath = doc.PathName;
            string fileName = Path.GetFileName(modelPath);
            string fileName2 = Path.GetFileNameWithoutExtension(modelPath);
            string folderPath = Path.GetDirectoryName(modelPath);
            string txtFile = folderPath + "\\" + fileName2 + ".txt";

            List<string> stringList = new List<string>();
            stringList.Add("Line 1");
            stringList.Add("Line 2");
            stringList.Add("Line 3");

            //using (StreamWriter writer = File.CreateText(txtFile))
            //{
            //    foreach(string curLine in stringList)
            //    {
            //        writer.WriteLine(curLine);
            //    }
            //}

            //if (File.Exists(txtFile))
            //{
            //    string[] textFile = File.ReadAllLines(txtFile);
            //}

            //FrmTestForm form1 = new FrmTestForm();
            //form1.Show();

            //TestData test1 = new TestData("This is a string", "this is another string", 10);

                return Result.Succeeded;


        }

    }
}