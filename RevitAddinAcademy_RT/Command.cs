#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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

            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select elements by rectangle:");
            List<CurveElement> curveList = new List<CurveElement>();

            WallType curWallType = GetWallTypeByName(doc, "Partition A30X");
            Level curLevel = GetLevelByName(doc, "_OVERALL FLOOR PLAN");

            MEPSystemType curSystemType = GetSystemTypeByName(doc, "Domestic Hot Water");
            PipeType curPipeType = GetPipeTypeByName(doc, "Default");

            using (Transaction t = new Transaction(doc))
            {
                t.Start();
                foreach (Element element in pickList)
                {
                    //Cast Element to something more specific
                    if (element is CurveElement)
                    {
                        CurveElement curve = (CurveElement)element;
                        //CurveElement curve2 = element as CurveElement;  //same thing, different syntax

                        curveList.Add(curve);

                        GraphicsStyle curGS = curve.LineStyle as GraphicsStyle;

                        switch (curGS.Name)
                        {
                            case "<Medium Lines>":
                                Debug.Print("found a medium line");
                                break;

                            case "<Thin Lines>":
                                Debug.Print("found a thin line");
                                break ;

                            case "<Wide Lines>":
                                Debug.Print("found a wide line");
                                break;

                            default:
                                Debug.Print("Found something else");
                                break;
                        }

                        Curve curCurve = curve.GeometryCurve;
                        XYZ startPoint = curCurve.GetEndPoint(0);
                        XYZ endPoint = curCurve.GetEndPoint(1);

                        //Wall newWall = Wall.Create(doc, curCurve, curWallType.Id, curLevel.Id, 15, 0, false, false);

                        Pipe newPipe = Pipe.Create(
                            doc,
                            curSystemType.Id,
                            curPipeType.Id,
                            curLevel.Id,
                            startPoint,
                            endPoint);


                        Debug.Print(curGS.Name);

                    }
                }
                t.Commit();
            }

                TaskDialog.Show("Complete", curveList.Count.ToString());

            return Result.Succeeded;
        }

        private WallType GetWallTypeByName(Document doc, string wallTypeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));

            foreach(Element curElem in collector)
            {
                WallType wallType = (WallType)curElem;

                if(wallType.Name == wallTypeName)
                    return wallType;
            }
            return null;
        }

        private Level GetLevelByName(Document doc, string levelName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Level));

            foreach (Element curElem in collector)
            {
                Level level = (Level)curElem;

                if (level.Name == levelName)
                    return level;
            }
            return null;
        }


        private MEPSystemType GetSystemTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(MEPSystemType));

            foreach (Element curElem in collector)
            {
                MEPSystemType curType = (MEPSystemType)curElem;

                if (curType.Name == typeName)
                    return curType;
            }
            return null;
        }

        private PipeType GetPipeTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(PipeType));

            foreach (Element curElem in collector)
            {
                PipeType curType = (PipeType)curElem;

                if (curType.Name == typeName)
                    return curType;
            }
            return null;
        }

    }
}
