using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;


namespace RevitAddinAcademy_RT
{
    public static class Utilities
    {
        //Get all rooms in the Revit model
        public static List<SpatialElement> GetAllRooms(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Rooms);

            List<SpatialElement> roomList = new List<SpatialElement>();

            foreach(Element e in collector)
            {
                SpatialElement curRoom = e as SpatialElement;
                roomList.Add(curRoom);
            }
            return roomList;
        }

        public static FamilySymbol GetFamilySymbolByName(Document doc, string familyName, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Family));

            foreach (Element element in collector)
            {
                Family curFamily = element as Family;

                if (curFamily.Name == familyName)
                {
                    ISet<ElementId> famSymbolIdList = curFamily.GetFamilySymbolIds();

                    foreach (ElementId famSymbolId in famSymbolIdList)
                    {
                        FamilySymbol curFS = doc.GetElement(famSymbolId) as FamilySymbol;

                        if (curFS.Name == typeName)
                            return curFS;
                    }
                }
            }

            return null;
        }

        //Get a Parameter Value where Value is string
        public static string GetParamValueAsString(Element curElem, string paramName)
        {
            foreach(Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                {
                    return curParam.AsString();
                }
            }
            return null;
        }
        
        //Get a Parameter Value where Value is double
        public static double GetParamValueAsDouble(Element curElem, string paramName)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                {
                    return curParam.AsDouble();
                }
            }
            return 0;
        }
       
        //Set Parameter value - string overload
        public static void SetParamValue(Element curElem, string paramName, string paramValue)
        {
            foreach(Parameter curParam in curElem.Parameters)
            {
                if(curParam.Definition.Name == paramName)
                {
                    curParam.Set(paramValue);
                }
            }
        }
        
        //Set Parameter value - double overload
        public static void SetParamValue(Element curElem, string paramName, double paramValue)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                {
                    curParam.Set(paramValue);
                }
            }
        }


        //needed for ReadExcel
        public struct DataSheet
        {
            public string Name;
            public List<string[]> DataList;

            public DataSheet(string name, List<string[]> dataList)
            {
                Name = name;
                DataList = dataList;
            }
        }

        public static List<string[]> GetDataBySheetName(List<DataSheet> dataSheets, string name)
        {
            foreach(DataSheet dataSheet in dataSheets)
            {
                if (dataSheet.Name == name)
                {
                    return dataSheet.DataList;
                }
            }
            return null;
        }

        //Read Excel data; all worksheets, all UsedRange
        public static List<DataSheet> ReadExcel(string filePath)
        {
            List<DataSheet> workbookData = new List<DataSheet>();

            try
            {
                //Setup Excel
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWb = excelApp.Workbooks.Open(filePath);
                Excel.Sheets excelWorksheets = excelWb.Sheets;

                foreach (Excel.Worksheet ws in excelWorksheets)
                {
                    string curName = ws.Name;
                    List<string[]> sheetData = new List<string[]>();

                    Excel.Range curRange = ws.UsedRange;
                    int rowCount = curRange.Rows.Count;
                    int columnCount = curRange.Columns.Count;

                    for (int i = 2; i <= rowCount; i++)
                    {
                        string[] curRow = new string[columnCount];
                        for (int j = 1; j <= columnCount; j++)
                        {
                            int index = j - 1;
                            Excel.Range curCell = ws.Cells[i, j];
                            curRow[index] = curCell.Value.ToString();
                        }
                        sheetData.Add(curRow);
                    }
                    DataSheet dataSheet = new DataSheet(curName, sheetData);
                    workbookData.Add(dataSheet);
                }

                excelWb.Close();
                excelApp.Quit();
            }
            catch (Exception e)
            {
                Debug.Print(e.Message);
            }

            return workbookData;
        }

    }
}
