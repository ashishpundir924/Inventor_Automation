using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Grid = System.Windows.Controls.Grid;
using TextBox = System.Windows.Controls.TextBox;

namespace DEFOR_Combinations
{
    public class SelectCombinationWindow : Window
    {
        private readonly ListBox _listBox;
        private readonly System.Collections.Generic.List<CombinationModel> _combinations;

        public int SelectedIndex { get; private set; } = -1;

        public SelectCombinationWindow(List<CombinationModel> combinations)
        {
            _combinations = combinations ?? new List<CombinationModel>();

            Title = "Select Combination";
            Width = 300;
            Height = 400;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;

            Grid root = new Grid { Margin = new Thickness(10) };
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            _listBox = new ListBox();
            foreach (CombinationModel combo in _combinations)
            {
                _listBox.Items.Add(combo.Name);
            }
            Grid.SetRow(_listBox, 0);
            root.Children.Add(_listBox);

            StackPanel buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 5, 0, 0)
            };

            Button okButton = new Button { Content = "OK", Width = 80, Margin = new Thickness(5) };
            okButton.Click += OkButton_Click;

            Button cancelButton = new Button { Content = "Cancel", Width = 80, Margin = new Thickness(5) };
            cancelButton.Click += (s, e) => DialogResult = false;

            buttons.Children.Add(okButton);
            buttons.Children.Add(cancelButton);

            Grid.SetRow(buttons, 1);
            root.Children.Add(buttons);

            Content = root;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (_listBox.SelectedIndex < 0)
            {
                MessageBox.Show("Please select a combination.", "Select Combination",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            SelectedIndex = _listBox.SelectedIndex;
            DialogResult = true;
        }
    }
}
