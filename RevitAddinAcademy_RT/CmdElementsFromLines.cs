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
    public class CmdElementsFromLines : IExternalCommand
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

            //Select elements
            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select elements by rectangle:");

            //Select Types
            Level curLevel = GetLevelByName(doc, "Level 1");
            WallType curWallType = GetWallTypeByName(doc, @"Generic - 8""");
            WallType curStorefrontType = GetWallTypeByName(doc, "Storefront");
            MEPSystemType curMechType = GetSystemTypeByName(doc, "Supply Air");
            MEPSystemType curPlumbType = GetSystemTypeByName(doc, "Domestic Hot Water");
            DuctType curDuctType = GetDuctTypeByName(doc, "Default");
            PipeType curPipeType = GetPipeTypeByName(doc, "Default");


            using (Transaction t = new Transaction(doc, "Elements from Lines"))
            {
                t.Start();
                foreach (Element element in pickList)
                {
                    if (element is CurveElement)
                    {
                        //Cast Element to CurveElement
                        CurveElement curve = (CurveElement)element;

                        //Get Geometry
                        Curve curCurve = curve.GeometryCurve;

                        //Switch by Line Style
                        GraphicsStyle curGS = curve.LineStyle as GraphicsStyle;

                        switch (curGS.Name)
                        {
                            case "A-GLAZ":
                                Wall newStorefront = Wall.Create(doc, curCurve, curStorefrontType.Id, curLevel.Id, 10, 0, false, false);
                                break;

                            case "A-WALL":
                                Wall newGeneric = Wall.Create(doc, curCurve, curWallType.Id, curLevel.Id, 10, 0, false, false);
                                break ;

                            case "M-DUCT":
                                XYZ startPoint = curCurve.GetEndPoint(0);
                                XYZ endPoint = curCurve.GetEndPoint(1);
                                Duct newDuct = Duct.Create(doc, curMechType.Id, curDuctType.Id, curLevel.Id, startPoint, endPoint);
                                break;

                            case "P-PIPE":
                                XYZ startPoint2 = curCurve.GetEndPoint(0);
                                XYZ endPoint2 = curCurve.GetEndPoint(1);
                                Pipe newPipe = Pipe.Create(doc, curPlumbType.Id, curPipeType.Id, curLevel.Id, startPoint2, endPoint2);
                                break;

                            default:
                                break;
                        }

                    }
                }
                t.Commit();
            }

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


        private DuctType GetDuctTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(DuctType));

            foreach (Element curElem in collector)
            {
                DuctType curType = (DuctType)curElem;

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
