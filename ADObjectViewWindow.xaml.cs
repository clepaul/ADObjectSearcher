/* Example DLL to extend the tool "Clipboard Accelerator"
Copyright (C) 2016 - 2020  Clemens Paul

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>. */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ADObjectSearcher
{
    /// <summary>
    /// Interaction logic for ADObjectViewWindow.xaml
    /// </summary>
    public partial class ADObjectViewWindow : Window
    {
        private static readonly List<ADObjectViewWindow> allInstancesOfThisClass = new List<ADObjectViewWindow>();

        public ADObjectViewWindow(List<ADObjectNameValuePair> adoNameValuePair)
        {
            InitializeComponent();

            dataGridADObjects.ItemsSource = adoNameValuePair;           

            ADObjectViewWindow.allInstancesOfThisClass.Add(this);

            // Make single cells selectable
            dataGridADObjects.SelectionMode = DataGridSelectionMode.Extended;
            dataGridADObjects.SelectionUnit = DataGridSelectionUnit.CellOrRowHeader;
        }


        // Source:
        // https://social.msdn.microsoft.com/Forums/vstudio/en-US/f9e984e3-a47b-42b0-b9bd-1f1c55e4de96/getting-all-running-instances-of-a-specific-class-type-using-reflection?forum=netfxbcl
        public static void CloseAllWindows()
        {
            try
            {
                allInstancesOfThisClass.ForEach(x => x.Close());
            }
            catch(Exception currEx)
            {
                MessageBox.Show(currEx.Message);
            }            
        }
       

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // Set the sorting direction
                // To be checked: The column shows it's sorted but actuall it is not
                // dataGridADObjects.Columns.First(x => x.Header.ToString() == "PropertyName").SortDirection = ListSortDirection.Ascending;                
            }
            catch (Exception curEx)
            {
            }
        }
    }
}
