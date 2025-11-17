using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.UI;
using Grid = System.Windows.Controls.Grid;
using TextBox = System.Windows.Controls.TextBox;

namespace DEFOR_Combinations
{
    public class EditCombinationWindow : Window
    {
        private readonly UIDocument _uidoc;
        private readonly ListBox _itemsList;
        private readonly TextBox _nameBox;

        public CombinationModel Combination { get; private set; }

        public EditCombinationWindow(UIDocument uidoc, CombinationModel existing = null)
        {
            _uidoc = uidoc;

            Title = existing == null ? "New Combination" : "Edit Combination";
            Width = 500;
            Height = 400;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;

            Grid root = new Grid { Margin = new Thickness(10) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // name
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // list
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // item buttons
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // ok/cancel

            // Name row
            StackPanel namePanel = new StackPanel { Orientation = Orientation.Horizontal };
            namePanel.Children.Add(new Label { Content = "Combination Name:", Margin = new Thickness(0, 0, 5, 0) });
            _nameBox = new TextBox { MinWidth = 200 };
            namePanel.Children.Add(_nameBox);
            Grid.SetRow(namePanel, 0);
            root.Children.Add(namePanel);

            // Items list
            _itemsList = new ListBox { Margin = new Thickness(0, 5, 0, 5) };
            Grid.SetRow(_itemsList, 1);
            root.Children.Add(_itemsList);

            // Item buttons
            StackPanel itemButtons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(0, 5, 0, 5)
            };

            Button addBtn = new Button { Content = "Add Item", Width = 90, Margin = new Thickness(5) };
            addBtn.Click += AddBtn_Click;

            Button editBtn = new Button { Content = "Edit Item", Width = 90, Margin = new Thickness(5) };
            editBtn.Click += EditBtn_Click;

            Button removeBtn = new Button { Content = "Remove Item", Width = 90, Margin = new Thickness(5) };
            removeBtn.Click += RemoveBtn_Click;

            itemButtons.Children.Add(addBtn);
            itemButtons.Children.Add(editBtn);
            itemButtons.Children.Add(removeBtn);

            Grid.SetRow(itemButtons, 2);
            root.Children.Add(itemButtons);

            // OK / Cancel buttons
            StackPanel okCancel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 5, 0, 0)
            };

            Button okButton = new Button { Content = "OK", Width = 80, Margin = new Thickness(5) };
            okButton.Click += OkButton_Click;

            Button cancelButton = new Button { Content = "Cancel", Width = 80, Margin = new Thickness(5) };
            cancelButton.Click += (s, e) => DialogResult = false;

            okCancel.Children.Add(okButton);
            okCancel.Children.Add(cancelButton);

            Grid.SetRow(okCancel, 3);
            root.Children.Add(okCancel);

            Content = root;

            if (existing == null)
            {
                Combination = new CombinationModel();
            }
            else
            {
                Combination = new CombinationModel
                {
                    Name = existing.Name,
                    Items = new List<CombinationItem>(existing.Items)
                };
            }

            _nameBox.Text = Combination.Name ?? string.Empty;
            RefreshItems();
        }

        private void RefreshItems()
        {
            _itemsList.Items.Clear();
            foreach (CombinationItem item in Combination.Items)
            {
                _itemsList.Items.Add(item);
            }
        }

        private void AddBtn_Click(object sender, RoutedEventArgs e)
        {
            EditItemWindow wnd = new EditItemWindow(_uidoc);
            bool? res = wnd.ShowDialog();
            if (res == true)
            {
                CombinationItem item = new CombinationItem
                {
                    FamilyName = wnd.FamilyName,
                    SymbolName = wnd.TypeName,
                    SymbolUniqueId = wnd.SymbolUniqueId,
                    Quantity = wnd.Quantity,
                    OffsetX = wnd.OffsetX,
                    OffsetY = wnd.OffsetY
                };
                Combination.Items.Add(item);
                RefreshItems();
            }
        }

        private void EditBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!(_itemsList.SelectedItem is CombinationItem selected))
            {
                MessageBox.Show("Please select an item to edit.", "Combination",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            EditItemWindow wnd = new EditItemWindow(_uidoc, selected);
            bool? res = wnd.ShowDialog();
            if (res == true)
            {
                selected.FamilyName = wnd.FamilyName;
                selected.SymbolName = wnd.TypeName;
                selected.SymbolUniqueId = wnd.SymbolUniqueId;
                selected.Quantity = wnd.Quantity;
                selected.OffsetX = wnd.OffsetX;
                selected.OffsetY = wnd.OffsetY;
                RefreshItems();
            }
        }

        private void RemoveBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!(_itemsList.SelectedItem is CombinationItem selected))
            {
                MessageBox.Show("Please select an item to remove.", "Combination",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            Combination.Items.Remove(selected);
            RefreshItems();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            string name = _nameBox.Text.Trim();
            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show("Combination name cannot be empty.", "Combination",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Combination.Name = name;
            DialogResult = true;
        }
    }
}
