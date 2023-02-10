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
    public class RAA_Session07Challenge : IExternalCommand
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

            frmToDo_Sln curForm = new frmToDo_Sln(doc.PathName);
            curForm.TopMost = true;
            curForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

            curForm.Show();

            return Result.Succeeded;


        }

    }
}