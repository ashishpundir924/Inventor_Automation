using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;

namespace DEFOR_Combinations
{
    [Transaction(TransactionMode.Manual)]
    public class PlaceCombinationCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<CombinationModel> combinations = CombinationStorage.Load();
            if (combinations == null || !combinations.Any())
            {
                TaskDialog.Show("Place Combination",
                    "There are no saved combinations to place. Use 'Manage Combinations' first.");
                return Result.Cancelled;
            }

            // Select combination
            SelectCombinationWindow selectWnd = new SelectCombinationWindow(combinations);
            bool? selRes = selectWnd.ShowDialog();
            if (selRes != true ||
                selectWnd.SelectedIndex < 0 ||
                selectWnd.SelectedIndex >= combinations.Count)
            {
                return Result.Cancelled;
            }

            CombinationModel combo = combinations[selectWnd.SelectedIndex];


            // 🚀 NEW LOGIC: USER SELECTS A FAMILY INSTANCE INSTEAD OF PICKING POINT
            Reference pickedRef = uidoc.Selection.PickObject(
                Autodesk.Revit.UI.Selection.ObjectType.Element,
                "Select a family instance to use as the base point."
            );

            Element pickedElement = doc.GetElement(pickedRef);
            FamilyInstance fi = pickedElement as FamilyInstance;

            if (fi == null)
            {
                TaskDialog.Show("Place Combination", "Please select a valid family instance.");
                return Result.Cancelled;
            }

            LocationPoint loc = fi.Location as LocationPoint;
            if (loc == null)
            {
                TaskDialog.Show("Place Combination", "Selected element does not have a point location.");
                return Result.Cancelled;
            }

            XYZ basePoint = loc.Point;  // ⭐ This is now the placement origin


            int totalPlaced = 0;
            List<string> missingSymbols = new List<string>();

            using (Transaction tx = new Transaction(doc, "Place Combination"))
            {
                tx.Start();

                foreach (CombinationItem item in combo.Items)
                {
                    Element el = null;
                    try
                    {
                        el = doc.GetElement(item.SymbolUniqueId);
                    }
                    catch { }

                    FamilySymbol symbol = el as FamilySymbol;
                    if (symbol == null)
                    {
                        missingSymbols.Add($"{item.FamilyName} : {item.SymbolName}");
                        continue;
                    }

                    if (!symbol.IsActive)
                    {
                        symbol.Activate();
                        doc.Regenerate();
                    }

                    XYZ offset = new XYZ(item.OffsetX, item.OffsetY, 0);

                    for (int i = 0; i < item.Quantity; i++)
                    {
                        XYZ pt = basePoint + offset;
                        try
                        {
                            doc.Create.NewFamilyInstance(pt, symbol, StructuralType.NonStructural);
                            totalPlaced++;
                        }
                        catch
                        {
                            // Skip failing instance
                        }
                    }
                }

                tx.Commit();
            }

            string report = $"Placed {totalPlaced} instances from combination '{combo.Name}'.";
            if (missingSymbols.Any())
            {
                report += "\n\nThe following symbols were missing:\n";
                foreach (string s in missingSymbols.Distinct())
                {
                    report += "- " + s + "\n";
                }
            }

            TaskDialog.Show("Place Combination", report);
            return Result.Succeeded;
        }
    }
}
