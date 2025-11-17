using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace DEFOR_Combinations
{
    [Transaction(TransactionMode.Manual)]
    public class ManageCombinationsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            List<CombinationModel> combinations = CombinationStorage.Load();
            bool done = false;

            while (!done)
            {
                TaskDialog dlg = new TaskDialog("Manage Combinations");
                dlg.MainInstruction = "Manage Combinations";
                dlg.MainContent = "Choose an action:";
                dlg.CommonButtons = TaskDialogCommonButtons.Close;
                dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Create New Combination");
                dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Edit Existing Combination");
                dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Delete Existing Combination");

                TaskDialogResult res = dlg.Show();
                if (res == TaskDialogResult.Close)
                {
                    done = true;
                    break;
                }

                if (res == TaskDialogResult.CommandLink1)
                {
                    // Create
                    EditCombinationWindow wnd = new EditCombinationWindow(uidoc);
                    bool? editRes = wnd.ShowDialog();
                    if (editRes == true)
                    {
                        CombinationModel newCombo = wnd.Combination;

                        bool nameExists = combinations.Any(c =>
                            string.Equals(c.Name, newCombo.Name, StringComparison.OrdinalIgnoreCase));

                        if (nameExists)
                        {
                            TaskDialog.Show("Manage Combinations",
                                "A combination with this name already exists. Please choose a different name.");
                        }
                        else
                        {
                            combinations.Add(newCombo);
                            CombinationStorage.Save(combinations);
                            TaskDialog.Show("Manage Combinations", "Combination created and saved.");
                        }
                    }
                }
                else if (res == TaskDialogResult.CommandLink2)
                {
                    // Edit
                    if (!combinations.Any())
                    {
                        TaskDialog.Show("Manage Combinations", "There are no combinations to edit.");
                        continue;
                    }

                    SelectCombinationWindow selectWnd = new SelectCombinationWindow(combinations);
                    bool? selRes = selectWnd.ShowDialog();
                    if (selRes == true && selectWnd.SelectedIndex >= 0 &&
                        selectWnd.SelectedIndex < combinations.Count)
                    {
                        CombinationModel existing = combinations[selectWnd.SelectedIndex];
                        EditCombinationWindow wnd = new EditCombinationWindow(uidoc, existing);
                        bool? editRes = wnd.ShowDialog();
                        if (editRes == true)
                        {
                            CombinationModel updated = wnd.Combination;

                            bool nameConflict = combinations
                                .Where((c, idx) => idx != selectWnd.SelectedIndex)
                                .Any(c => string.Equals(c.Name, updated.Name, StringComparison.OrdinalIgnoreCase));

                            if (nameConflict)
                            {
                                TaskDialog.Show("Manage Combinations",
                                    "Another combination already uses this name. Please choose a different name.");
                            }
                            else
                            {
                                combinations[selectWnd.SelectedIndex] = updated;
                                CombinationStorage.Save(combinations);
                                TaskDialog.Show("Manage Combinations", "Combination updated and saved.");
                            }
                        }
                    }
                }
                else if (res == TaskDialogResult.CommandLink3)
                {
                    // Delete
                    if (!combinations.Any())
                    {
                        TaskDialog.Show("Manage Combinations", "There are no combinations to delete.");
                        continue;
                    }

                    SelectCombinationWindow selectWnd = new SelectCombinationWindow(combinations);
                    bool? selRes = selectWnd.ShowDialog();
                    if (selRes == true && selectWnd.SelectedIndex >= 0 &&
                        selectWnd.SelectedIndex < combinations.Count)
                    {
                        CombinationModel toDelete = combinations[selectWnd.SelectedIndex];

                        TaskDialog confirm = new TaskDialog("Delete Combination");
                        confirm.MainInstruction = "Delete Combination";
                        confirm.MainContent = $"Are you sure you want to delete '{toDelete.Name}'?";
                        confirm.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;

                        TaskDialogResult cRes = confirm.Show();
                        if (cRes == TaskDialogResult.Yes)
                        {
                            combinations.RemoveAt(selectWnd.SelectedIndex);
                            CombinationStorage.Save(combinations);
                            TaskDialog.Show("Manage Combinations", "Combination deleted.");
                        }
                    }
                }
            }

            return Result.Succeeded;
        }
    }
}
