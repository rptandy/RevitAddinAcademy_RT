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

#endregion

namespace RevitAddinAcademy_RT
{
    [Transaction(TransactionMode.Manual)]
    public class Command01Challenge : IExternalCommand
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

            //Set offset value for each successive text note
            double offset = 0.05;
            double offsetCalc = offset * doc.ActiveView.Scale;

            //Create text insertion points
            XYZ curPoint = new XYZ(0, 0, 0);
            XYZ offsetPoint = new XYZ(0, offsetCalc, 0);

            //Get types of text notes
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(TextNoteType));

            //Create transaction
            Transaction t = new Transaction(doc, "Play Fizzbuzz");
            t.Start();

            //For loop
            int range = 100;
            for (int i = 1; i <= range; i++)
            {
                //Calculate modulus 
                int a = i % 3;
                int b = i % 5;
                string text = "";

                //Divisible by both 3 and 5
                if (a == 0 && b == 0)
                {
                    text = "FIZZBUZZ";
                }
                //Divisible by 3
                else if (a == 0 && b > 0)
                {
                    text = "FIZZ";
                }
                //Divisible by 5
                else if (a > 0 && b == 0)
                {
                    text = "BUZZ";
                }
                //Divisible by neither 3 or 5
                else
                {
                    text = i.ToString();
                }

                TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, text, collector.FirstElementId());
                curPoint = curPoint.Subtract(offsetPoint);
            }

            //Finish transaction
            t.Commit();
            t.Dispose();

            return Result.Succeeded;
        }
    }
}
