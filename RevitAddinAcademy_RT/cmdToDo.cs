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
using Forms = System.Windows.Forms;

#endregion

namespace RevitAddinAcademy_RT
{
    [Transaction(TransactionMode.Manual)]
    public class CmdToDo : IExternalCommand
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

            //get filepath
            string modelPath = doc.PathName;
            string fileName = Path.GetFileNameWithoutExtension(modelPath);
            string folderPath = Path.GetDirectoryName(modelPath);
            string txtFile = folderPath + "\\" + fileName + ".txt";

            //Launch form
            FrmToDo form1 = new FrmToDo(txtFile);
            form1.FormClosed += (sender, e) =>
            {
                List<string> toDoList = form1.ToDoList;

                if (toDoList != null)
                {
                    using (StreamWriter writer = File.CreateText(txtFile))
                    {
                        foreach (string toDo in toDoList)
                        {
                            writer.WriteLine(toDo);
                        }
                    }
                }
            };

            form1.Show();

            return Result.Succeeded;


        }

    }
}