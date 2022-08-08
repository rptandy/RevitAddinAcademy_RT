#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.IO;

#endregion

namespace RevitAddinAcademy_RT
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            //step 1: create ribbon tab
            try
            {
                a.CreateRibbonTab("Revit Add-in Academy");
            }
            catch (Exception)
            {
                Debug.Print("Tab already exists");
            }

            //step 2: create ribbon panel
            RibbonPanel curPanel = CreateRibbonPanel(a, "Revit Add-in Academy", "Revit Tools");

            //step 3: create button data instances
            PushButtonData pData1 = new PushButtonData("tool1", "Project Setup", GetAssemblyName(), "RevitAddinAcademy_RT.CmdProjectSetup");
            PushButtonData pData2 = new PushButtonData("tool2", "Delete Backups", GetAssemblyName(), "RevitAddinAcademy_RT.CmdDeleteBackups");
            PushButtonData pData3 = new PushButtonData("tool3", "Tool 3", GetAssemblyName(), "RevitAddinAcademy_RT.Command");
            PushButtonData pData4 = new PushButtonData("tool4", "Tool 4", GetAssemblyName(), "RevitAddinAcademy_RT.Command");
            PushButtonData pData5 = new PushButtonData("tool5", "Tool 5", GetAssemblyName(), "RevitAddinAcademy_RT.Command");
            PushButtonData pData6 = new PushButtonData("tool6", "Elements from Lines", GetAssemblyName(), "RevitAddinAcademy_RT.CmdElementsFromLines");
            PushButtonData pData7 = new PushButtonData("tool7", "Add Furniture", GetAssemblyName(), "RevitAddinAcademy_RT.CmdMovingDay");
            PushButtonData pData8 = new PushButtonData("tool8", "Tool 8", GetAssemblyName(), "RevitAddinAcademy_RT.Command");
            PushButtonData pData9 = new PushButtonData("tool9", "Tool 9", GetAssemblyName(), "RevitAddinAcademy_RT.Command");
            PushButtonData pData10 = new PushButtonData("tool10", "Tool 10", GetAssemblyName(), "RevitAddinAcademy_RT.Command");

            SplitButtonData sData1 = new SplitButtonData("splitButton1", "Split Button 1");

            PulldownButtonData pbData1 = new PulldownButtonData("pulldownButton1", "More Tools");

            //step 4: add images
            pData1.Image = BitmapToImagesource(RevitAddinAcademy_RT.Properties.Resources.Blue_16);
            pData1.LargeImage = BitmapToImagesource(RevitAddinAcademy_RT.Properties.Resources.Blue_32);

            pData2.Image = BitmapToImagesource(RevitAddinAcademy_RT.Properties.Resources.Green_16);
            pData2.LargeImage = BitmapToImagesource(RevitAddinAcademy_RT.Properties.Resources.Green_32);

            pData3.Image = BitmapToImagesource(RevitAddinAcademy_RT.Properties.Resources.Blue_16);
            pData3.LargeImage = BitmapToImagesource(RevitAddinAcademy_RT.Properties.Resources.Blue_32);

            pData4.Image = BitmapToImagesource(RevitAddinAcademy_RT.Properties.Resources.Yellow_16);
            pData4.LargeImage = BitmapToImagesource(RevitAddinAcademy_RT.Properties.Resources.Yellow_32);

            pData5.Image = BitmapToImagesource(RevitAddinAcademy_RT.Properties.Resources.Green_16);
            pData5.LargeImage = BitmapToImagesource(RevitAddinAcademy_RT.Properties.Resources.Green_32);

            pData6.Image = BitmapToImagesource(RevitAddinAcademy_RT.Properties.Resources.Red_16);
            pData6.LargeImage = BitmapToImagesource(RevitAddinAcademy_RT.Properties.Resources.Red_32);

            pData7.Image = BitmapToImagesource(RevitAddinAcademy_RT.Properties.Resources.Yellow_16);
            pData7.LargeImage = BitmapToImagesource(RevitAddinAcademy_RT.Properties.Resources.Yellow_32);

            pbData1.Image = BitmapToImagesource(RevitAddinAcademy_RT.Properties.Resources.Red_16);
            pbData1.LargeImage = BitmapToImagesource(RevitAddinAcademy_RT.Properties.Resources.Red_32);

            //step 5: add tooltip info
            pData1.ToolTip = "Create levels, views, and sheets from Excel data";
            pData2.ToolTip = "Delete all Revit backup files in a directory";
            pData3.ToolTip = "Tool 3 tooltip";
            pData4.ToolTip = "Tool 4 tooltip";
            pData5.ToolTip = "Tool 5 tooltip";
            pData6.ToolTip = "Create Revit elements from lines";
            pData7.ToolTip = "Add furniture to rooms from Excel data";
            pData8.ToolTip = "Tool 8 tooltip";
            pData9.ToolTip = "Tool 9 tooltip";
            pData10.ToolTip = "Tool 10 tooltip";

            //step 6: create buttons
            //push buttons (1-2)
            PushButton tool1 = curPanel.AddItem(pData1) as PushButton;
            PushButton tool2 = curPanel.AddItem(pData2) as PushButton;

            //stacked row (3-5)
            curPanel.AddStackedItems(pData3, pData4, pData5);

            //split button (6-7)
            SplitButton splitButton1 = curPanel.AddItem(sData1) as SplitButton;
            splitButton1.AddPushButton(pData6);
            splitButton1.AddPushButton(pData7);

            //pulldown (8-10)
            PulldownButton pulldownButton1 = curPanel.AddItem(pbData1) as PulldownButton;
            pulldownButton1.AddPushButton(pData8);
            pulldownButton1.AddPushButton(pData9);
            pulldownButton1.AddPushButton(pData10);

            return Result.Succeeded;
        }

        private RibbonPanel CreateRibbonPanel(UIControlledApplication a, string tabName, string panelName)
        {
            foreach(RibbonPanel tmpPanel in a.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                    return tmpPanel;
            }

            RibbonPanel returnPanel = a.CreateRibbonPanel(tabName, panelName);
            return returnPanel;
        }

        private string GetAssemblyName()
        {
            return Assembly.GetExecutingAssembly().Location;
        }

        private BitmapImage BitmapToImagesource(System.Drawing.Bitmap bm)
        {
            using(MemoryStream ms = new MemoryStream())
            {
                bm.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                ms.Position = 0;
                BitmapImage bmi = new BitmapImage();
                bmi.BeginInit();
                bmi.StreamSource = ms;
                bmi.CacheOption = BitmapCacheOption.OnLoad;
                bmi.EndInit();
                return bmi;
            }
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
