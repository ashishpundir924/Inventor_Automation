using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Grid = System.Windows.Controls.Grid;
using TextBox = System.Windows.Controls.TextBox;

namespace DEFOR_Combinations
{
    public class EditItemWindow : Window
    {
        private readonly UIDocument _uidoc;

        private TextBox _familyNameBox;
        private TextBox _typeNameBox;
        private TextBox _quantityBox;
        private TextBox _offsetXBox;
        private TextBox _offsetYBox;

        public string FamilyName { get; private set; }
        public string TypeName { get; private set; }
        public string SymbolUniqueId { get; private set; }
        public int Quantity { get; private set; }
        public double OffsetX { get; private set; }
        public double OffsetY { get; private set; }

        public EditItemWindow(UIDocument uidoc, CombinationItem existingItem = null)
        {
            _uidoc = uidoc;

            Title = existingItem == null ? "Add Item" : "Edit Item";
            Width = 420;
            Height = 280;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;

            Grid root = new Grid { Margin = new Thickness(10) };

            for (int i = 0; i < 6; i++)
            {
                root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            }
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // buttons

            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            // Row 0: Family + Pick
            Label familyLabel = new Label { Content = "Family:", Margin = new Thickness(0, 0, 5, 0) };
            Grid.SetRow(familyLabel, 0);
            Grid.SetColumn(familyLabel, 0);
            root.Children.Add(familyLabel);

            _familyNameBox = new TextBox { IsReadOnly = true, Margin = new Thickness(2) };
            Grid.SetRow(_familyNameBox, 0);
            Grid.SetColumn(_familyNameBox, 1);
            root.Children.Add(_familyNameBox);

            Button pickButton = new Button { Content = "Pick From Project", Margin = new Thickness(2) };
            pickButton.Click += PickButton_Click;
            Grid.SetRow(pickButton, 0);
            Grid.SetColumn(pickButton, 2);
            root.Children.Add(pickButton);

            // Row 1: Type name
            Label typeLabel = new Label { Content = "Type:", Margin = new Thickness(0, 5, 5, 0) };
            Grid.SetRow(typeLabel, 1);
            Grid.SetColumn(typeLabel, 0);
            root.Children.Add(typeLabel);

            _typeNameBox = new TextBox { IsReadOnly = true, Margin = new Thickness(2) };
            Grid.SetRow(_typeNameBox, 1);
            Grid.SetColumn(_typeNameBox, 1);
            Grid.SetColumnSpan(_typeNameBox, 2);
            root.Children.Add(_typeNameBox);

            // Row 2: Quantity
            Label qtyLabel = new Label { Content = "Quantity:", Margin = new Thickness(0, 5, 5, 0) };
            Grid.SetRow(qtyLabel, 2);
            Grid.SetColumn(qtyLabel, 0);
            root.Children.Add(qtyLabel);

            _quantityBox = new TextBox { Margin = new Thickness(2), Text = "1" };
            Grid.SetRow(_quantityBox, 2);
            Grid.SetColumn(_quantityBox, 1);
            Grid.SetColumnSpan(_quantityBox, 2);
            root.Children.Add(_quantityBox);

            // Row 3: Offset X
            Label offsetXLabel = new Label { Content = "Offset X:", Margin = new Thickness(0, 5, 5, 0) };
            Grid.SetRow(offsetXLabel, 3);
            Grid.SetColumn(offsetXLabel, 0);
            root.Children.Add(offsetXLabel);

            _offsetXBox = new TextBox { Margin = new Thickness(2), Text = "0" };
            Grid.SetRow(_offsetXBox, 3);
            Grid.SetColumn(_offsetXBox, 1);
            Grid.SetColumnSpan(_offsetXBox, 2);
            root.Children.Add(_offsetXBox);

            // Row 4: Offset Y
            Label offsetYLabel = new Label { Content = "Offset Y:", Margin = new Thickness(0, 5, 5, 0) };
            Grid.SetRow(offsetYLabel, 4);
            Grid.SetColumn(offsetYLabel, 0);
            root.Children.Add(offsetYLabel);

            _offsetYBox = new TextBox { Margin = new Thickness(2), Text = "0" };
            Grid.SetRow(_offsetYBox, 4);
            Grid.SetColumn(_offsetYBox, 1);
            Grid.SetColumnSpan(_offsetYBox, 2);
            root.Children.Add(_offsetYBox);

            // Spacer row 5
            root.RowDefinitions[5].Height = new GridLength(1, GridUnitType.Star);

            // Buttons row
            StackPanel buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(5)
            };

            Button okButton = new Button { Content = "OK", Width = 80, Margin = new Thickness(5) };
            okButton.Click += OkButton_Click;

            Button cancelButton = new Button { Content = "Cancel", Width = 80, Margin = new Thickness(5) };
            cancelButton.Click += (s, e) => DialogResult = false;

            buttons.Children.Add(okButton);
            buttons.Children.Add(cancelButton);

            Grid.SetRow(buttons, 6);
            Grid.SetColumn(buttons, 0);
            Grid.SetColumnSpan(buttons, 3);
            root.Children.Add(buttons);

            Content = root;

            if (existingItem != null)
            {
                FamilyName = existingItem.FamilyName;
                TypeName = existingItem.SymbolName;
                SymbolUniqueId = existingItem.SymbolUniqueId;
                Quantity = existingItem.Quantity;
                OffsetX = existingItem.OffsetX;
                OffsetY = existingItem.OffsetY;

                _familyNameBox.Text = FamilyName;
                _typeNameBox.Text = TypeName;
                _quantityBox.Text = Quantity.ToString(CultureInfo.InvariantCulture);
                _offsetXBox.Text = OffsetX.ToString(CultureInfo.InvariantCulture);
                _offsetYBox.Text = OffsetY.ToString(CultureInfo.InvariantCulture);
            }
        }

        private void PickButton_Click(object sender, RoutedEventArgs e)
        {
            Document doc = _uidoc.Document;
            SelectFamilyTypeWindow selector = new SelectFamilyTypeWindow(doc);
            bool? res = selector.ShowDialog();
            if (res == true)
            {
                FamilyName = selector.SelectedFamilyName;
                TypeName = selector.SelectedSymbolName;
                SymbolUniqueId = selector.SelectedSymbolUniqueId;

                _familyNameBox.Text = FamilyName;
                _typeNameBox.Text = TypeName;
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(FamilyName) ||
                string.IsNullOrWhiteSpace(TypeName) ||
                string.IsNullOrWhiteSpace(SymbolUniqueId))
            {
                MessageBox.Show("Please pick a family type from the project.", "Item",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!int.TryParse(_quantityBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int qty) || qty <= 0)
            {
                MessageBox.Show("Quantity must be a positive integer.", "Item",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!double.TryParse(_offsetXBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double dx))
            {
                MessageBox.Show("Offset X must be a number.", "Item",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!double.TryParse(_offsetYBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double dy))
            {
                MessageBox.Show("Offset Y must be a number.", "Item",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Quantity = qty;
            OffsetX = dx;
            OffsetY = dy;

            DialogResult = true;
        }
    }
}
