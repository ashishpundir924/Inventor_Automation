using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Grid = System.Windows.Controls.Grid;
using TextBox = System.Windows.Controls.TextBox;

namespace DEFOR_Combinations
{
    public class SelectFamilyTypeWindow : Window
    {
        private readonly Document _doc;
        private readonly TreeView _treeView;
        private readonly Button _okButton;
        private readonly Button _cancelButton;

        public string SelectedFamilyName { get; private set; }
        public string SelectedSymbolName { get; private set; }
        public string SelectedSymbolUniqueId { get; private set; }

        public SelectFamilyTypeWindow(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));

            Title = "Select Family Type";
            Width = 400;
            Height = 500;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;

            Grid root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            _treeView = new TreeView();
            Grid.SetRow(_treeView, 0);
            root.Children.Add(_treeView);

            StackPanel buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(5)
            };

            _okButton = new Button { Content = "OK", Width = 80, Margin = new Thickness(5) };
            _okButton.Click += OkButton_Click;

            _cancelButton = new Button { Content = "Cancel", Width = 80, Margin = new Thickness(5) };
            _cancelButton.Click += (s, e) => DialogResult = false;

            buttonPanel.Children.Add(_okButton);
            buttonPanel.Children.Add(_cancelButton);

            Grid.SetRow(buttonPanel, 1);
            root.Children.Add(buttonPanel);

            Content = root;

            LoadFamilies();
        }

        private void LoadFamilies()
        {
            var symbols = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .ToList();

            var grouped = symbols
                .GroupBy(s => s.Family.Name)
                .OrderBy(g => g.Key);

            foreach (var familyGroup in grouped)
            {
                TreeViewItem familyItem = new TreeViewItem
                {
                    Header = familyGroup.Key,
                    IsExpanded = false
                };

                foreach (var symbol in familyGroup.OrderBy(s => s.Name))
                {
                    TreeViewItem symbolItem = new TreeViewItem
                    {
                        Header = symbol.Name,
                        Tag = symbol
                    };
                    familyItem.Items.Add(symbolItem);
                }

                _treeView.Items.Add(familyItem);
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            TreeViewItem selectedItem = _treeView.SelectedItem as TreeViewItem;
            if (selectedItem == null || !(selectedItem.Tag is FamilySymbol symbol))
            {
                MessageBox.Show("Please select a specific family type.", "Select Family Type",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            SelectedFamilyName = symbol.Family.Name;
            SelectedSymbolName = symbol.Name;
            SelectedSymbolUniqueId = symbol.UniqueId;

            DialogResult = true;
        }
    }
}
