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
using Forms = System.Windows.Forms;

#endregion

namespace RevitAddinAcademy_RT
{
    [Transaction(TransactionMode.Manual)]
    public class CmdMovingDay : IExternalCommand
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

            Forms.OpenFileDialog openFile = new Forms.OpenFileDialog();
            openFile.InitialDirectory = "C:\\";
            openFile.Multiselect = false;
            openFile.Filter = "Excel files | *.xls; *.xlsx; *.xlsm | All files |*.*";

            if (openFile.ShowDialog() == Forms.DialogResult.OK)
            {
                try
                {
                    List<Utilities.DataSheet> allData = Utilities.ReadExcel(openFile.FileName);

                    List<string[]> furnitureData = Utilities.GetDataBySheetName(allData, "Furniture types");
                    List<string[]> setData = Utilities.GetDataBySheetName(allData, "Furniture sets");

                    List<Furniture> inventory = new List<Furniture>();
                    foreach (string[] f in furnitureData)
                    {
                        Furniture curFurniture = new Furniture(f[0], f[1], f[2]);
                        inventory.Add(curFurniture);
                    }

                    List<FurnitureSet> furnitureSets = new List<FurnitureSet>();
                    foreach (string[] s in setData)
                    {
                        FurnitureSet curSet = new FurnitureSet(s[0], s[1], s[2], inventory);
                        furnitureSets.Add(curSet);
                    }

                    using (Transaction t = new Transaction(doc))
                    {
                        t.Start("Moving Day");

                        List<SpatialElement> rooms = Utilities.GetAllRooms(doc);

                        foreach (SpatialElement room in rooms)
                        {
                            string curSetName = Utilities.GetParamValueAsString(room, "Furniture Set");
                            FurnitureSet curSet = FurnitureSet.GetFurnitureSetByName(furnitureSets, curSetName);

                            if (curSet != null)
                            {
                                Utilities.SetParamValue(room, "Furniture Count", curSet.FurnitureCount);

                                LocationPoint roomLocation = room.Location as LocationPoint;
                                XYZ roomPoint = roomLocation.Point;

                                foreach (Furniture furniture in curSet.FurnitureList)
                                {
                                    string name = furniture.Name;
                                    FamilySymbol curFS = Utilities.GetFamilySymbolByName(doc, furniture.FamilyName, furniture.FamilyType);
                                    if (curFS != null)
                                    {
                                        curFS.Activate();
                                        FamilyInstance curFI = doc.Create.NewFamilyInstance(roomPoint, curFS, StructuralType.NonStructural);
                                    }
                                }

                            }
                        }

                        t.Commit();
                    }
                }
                catch (Exception e)
                {
                    Debug.Print(e.Message);
                }
                
            }

            return Result.Succeeded;


        }

    }
}