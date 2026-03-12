using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ColumnDimensioner
{
    public partial class ColumnDimensionerWPF : Window
    {
        public string DimensionColumnsButtonName;
        public DimensionType SelectedDimensionType;

        public bool IndentationFirstRowDimensionsIsChecked;
        public string IndentationFirstRowDimensions;

        public bool IndentationSecondRowDimensionsIsChecked;
        public string IndentationSecondRowDimensions;
        ColumnDimensionerSettings ColumnDimensionerSettingsItem = null;

        public ColumnDimensionerWPF(List<DimensionType> dimensionTypesList)
        {
            ColumnDimensionerSettingsItem = new ColumnDimensionerSettings().GetSettings();
            InitializeComponent();
            comboBox_DimensionType.ItemsSource = dimensionTypesList;
            comboBox_DimensionType.DisplayMemberPath = "Name";

            if (ColumnDimensionerSettingsItem != null)
            {
                if (ColumnDimensionerSettingsItem.DimensionColumnsButtonName == "radioButton_VisibleInView")
                {
                    radioButton_VisibleInView.IsChecked = true;
                }
                else
                {
                    radioButton_Selected.IsChecked = true;
                }

                DimensionType savedType = dimensionTypesList.FirstOrDefault(dt => dt.Name == ColumnDimensionerSettingsItem.SelectedDimensionTypeName);
                comboBox_DimensionType.SelectedItem = savedType ?? comboBox_DimensionType.Items[0];

                checkBox_IndentationFirstRowDimensions.IsChecked = ColumnDimensionerSettingsItem.IndentationFirstRowDimensionsIsChecked;
                textBox_IndentationFirstRowDimensions.Text = ColumnDimensionerSettingsItem.IndentationFirstRowDimensions;

                checkBox_IndentationSecondRowDimensions.IsChecked = ColumnDimensionerSettingsItem.IndentationSecondRowDimensionsIsChecked;
                textBox_IndentationSecondRowDimensions.Text = ColumnDimensionerSettingsItem.IndentationSecondRowDimensions;
            }
            else
            {
                comboBox_DimensionType.SelectedItem = comboBox_DimensionType.Items[0];
                checkBox_IndentationFirstRowDimensions.IsChecked = true;
                textBox_IndentationFirstRowDimensions.Text = "700";

                checkBox_IndentationSecondRowDimensions.IsChecked = true;
                textBox_IndentationSecondRowDimensions.Text = "1400";
            }
        }

        private void btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            if (!SaveSettings())
            {
                return;
            }

            DialogResult = true;
            Close();
        }

        private void ColumnDimensionerWPF_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter || e.Key == Key.Space)
            {
                if (!SaveSettings())
                {
                    return;
                }

                DialogResult = true;
                Close();
            }
            else if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }

        private void btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private bool SaveSettings()
        {
            ColumnDimensionerSettingsItem = new ColumnDimensionerSettings();
            System.Windows.Controls.Grid dimensionColumnsGrid = groupBox_DimensionColumns.Content as System.Windows.Controls.Grid;
            DimensionColumnsButtonName = dimensionColumnsGrid
                ?.Children.OfType<RadioButton>()
                .FirstOrDefault(rb => rb.IsChecked == true)?.Name ?? "radioButton_VisibleInView";
            ColumnDimensionerSettingsItem.DimensionColumnsButtonName = DimensionColumnsButtonName;

            SelectedDimensionType = comboBox_DimensionType.SelectedItem as DimensionType
                ?? comboBox_DimensionType.Items.Cast<DimensionType>().FirstOrDefault();

            if (SelectedDimensionType == null)
            {
                MessageBox.Show("Не удалось определить тип размера. Выберите тип и повторите.", "ColumnDimensioner", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            ColumnDimensionerSettingsItem.SelectedDimensionTypeName = SelectedDimensionType.Name;

            IndentationFirstRowDimensionsIsChecked = checkBox_IndentationFirstRowDimensions.IsChecked == true;
            ColumnDimensionerSettingsItem.IndentationFirstRowDimensionsIsChecked = IndentationFirstRowDimensionsIsChecked;

            IndentationFirstRowDimensions = textBox_IndentationFirstRowDimensions.Text;
            ColumnDimensionerSettingsItem.IndentationFirstRowDimensions = IndentationFirstRowDimensions;

            IndentationSecondRowDimensionsIsChecked = checkBox_IndentationSecondRowDimensions.IsChecked == true;
            ColumnDimensionerSettingsItem.IndentationSecondRowDimensionsIsChecked = IndentationSecondRowDimensionsIsChecked;

            IndentationSecondRowDimensions = textBox_IndentationSecondRowDimensions.Text;
            ColumnDimensionerSettingsItem.IndentationSecondRowDimensions = IndentationSecondRowDimensions;
            ColumnDimensionerSettingsItem.SaveSettings();

            return true;
        }

        private void checkBox_IndentationFirstRowDimensions_Checked(object sender, RoutedEventArgs e)
        {
            bool isEnabled = checkBox_IndentationFirstRowDimensions.IsChecked == true;
            label_IndentationFirstRowDimensions.IsEnabled = isEnabled;
            textBox_IndentationFirstRowDimensions.IsEnabled = isEnabled;
            label_IndentationFirstRowDimensionsMM.IsEnabled = isEnabled;
        }

        private void checkBox_IndentationSecondRowDimensions_Checked(object sender, RoutedEventArgs e)
        {
            bool isEnabled = checkBox_IndentationSecondRowDimensions.IsChecked == true;
            label_IndentationSecondRowDimensions.IsEnabled = isEnabled;
            textBox_IndentationSecondRowDimensions.IsEnabled = isEnabled;
            label_IndentationSecondRowDimensionsMM.IsEnabled = isEnabled;
        }
    }
}
