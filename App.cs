using System;
using System.Reflection;
using Autodesk.Revit.UI;

namespace DEFOR_Combinations
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            // Ribbon tab
            string tabName = "DEFOR Tool";
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch
            {
                // Tab may already exist – ignore
            }

            // Panel
            string panelName = "Combinations";
            RibbonPanel panel = null;
            foreach (RibbonPanel p in application.GetRibbonPanels(tabName))
            {
                if (p.Name == panelName)
                {
                    panel = p;
                    break;
                }
            }

            if (panel == null)
            {
                panel = application.CreateRibbonPanel(tabName, panelName);
            }

            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            // Manage Combinations button
            PushButtonData manageData = new PushButtonData(
                "ManageCombinations",
                "Manage\nCombinations",
                assemblyPath,
                "DEFOR_Combinations.ManageCombinationsCommand");

            PushButton manageBtn = panel.AddItem(manageData) as PushButton;
            if (manageBtn != null)
            {
                manageBtn.ToolTip = "Create, edit, and delete saved combinations of family types.";
            }

            // Place Combination button
            PushButtonData placeData = new PushButtonData(
                "PlaceCombination",
                "Place\nCombination",
                assemblyPath,
                "DEFOR_Combinations.PlaceCombinationCommand");

            PushButton placeBtn = panel.AddItem(placeData) as PushButton;
            if (placeBtn != null)
            {
                placeBtn.ToolTip = "Place a saved combination of family instances into the model.";
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
