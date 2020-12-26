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
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System;
using System.Linq;

namespace ADObjectSearcher
{
    /// <summary>
    /// Interaction logic for ADObjectSearcherUI.xaml
    /// </summary>
    public partial class ADObjectSearcherWindow : Window
    {
        //ObservableCollection<MainDataGridEntry> gridDataEntries = new ObservableCollection<MainDataGridEntry>();
        List<MainDataGridEntry> gridDataEntries = new List<MainDataGridEntry>();
        List<MainDataGridEntry> gridDataFoundEntries = new List<MainDataGridEntry>();

        public ADObjectSearcherWindow(string[] ClipboardText, XMLRecord xmlRecord)
        {
            InitializeComponent();

            textBlockWildcardWarning.Visibility = Visibility.Hidden;

            // Put the domain name into the textbox
            textBoxDomainName.Text = SearchADO.GetCurrentDomain();

            foreach (string cs in ClipboardText)
            {
                //MainDataGridEntry tempData = new MainDataGridEntry { LookupName = cs, LdapQueryString = $"(samaccountname={cs})" };
                MainDataGridEntry tempData = new MainDataGridEntry { LookupName = cs, LdapQueryString = $"(|(sAMAccountName={cs})(cn={cs})(mail={cs}))" };                
                gridDataEntries.Add(tempData);
            }

            dataGrid.ItemsSource = gridDataEntries;

            dataGrid.SelectionMode = DataGridSelectionMode.Extended;
            dataGrid.SelectionUnit = DataGridSelectionUnit.CellOrRowHeader;
        }


        private void checkBoxExactSearch_Click(object sender, RoutedEventArgs e)
        {
            // Set the dataGrid to its initial data source and refresh the datagrid to show the new source
            // This is required because after the search the dataGrid source is set to the "gridDataFoundEntries" object
            dataGrid.ItemsSource = gridDataEntries;
            dataGrid.Items.Refresh();
            dataGrid.Columns.First(x => x.Header.ToString() == "ADObject").Visibility = Visibility.Hidden;


            if (checkBoxExactSearch.IsChecked == false)
            {
                textBlockWildcardWarning.Visibility = Visibility.Visible;

                foreach (MainDataGridEntry dd in gridDataEntries)
                {
                    //dd.LdapQueryString = $"(samaccountname=*{dd.LookupName}*)";                    
                    dd.LdapQueryString = $"(|(sAMAccountName=*{dd.LookupName}*)(cn=*{dd.LookupName}*)(mail=*{dd.LookupName}*))";
                }
            }
            else if (checkBoxExactSearch.IsChecked == true)
            {
                textBlockWildcardWarning.Visibility = Visibility.Hidden;

                foreach (MainDataGridEntry dd in gridDataEntries)
                {
                    //dd.LdapQueryString = $"(samaccountname={dd.LookupName})";
                    dd.LdapQueryString = $"(|(sAMAccountName={dd.LookupName})(cn={dd.LookupName})(mail={dd.LookupName}))";
                }
            }

            dataGrid.Items.Refresh();
        }


        // Source: https://stackoverflow.com/questions/22790181/wpf-datagrid-row-double-click-event-programmatically
        private void DataGridRow_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            DataGridRow row = sender as DataGridRow;
            MainDataGridEntry dgridentry = row.Item as MainDataGridEntry;

            foreach (List<ADObjectNameValuePair> adopair in dgridentry.ADObject.TailoredADObjects)
            {
                ADObjectViewWindow adObjectsWindow = new ADObjectViewWindow(adopair);
                adObjectsWindow.Show();
                adObjectsWindow.Title = adopair.First(x => x.PropertyName == "distinguishedname").PropertyValue;
            }      
        }


        private async void buttonSearchAD_Click(object sender, RoutedEventArgs e)
        {
            // (re-)set the initial datagrid entries. This is required if another search is triggered 
            // and the ItemsSource has already been set to the List containing the found values which happens below in the function
            dataGrid.ItemsSource = gridDataEntries;

            dataGrid.Columns.First(x => x.Header.ToString() == "ADObject").Visibility = Visibility.Hidden;

            // Clear the "found AD Objects" list to prevent duplicates
            gridDataFoundEntries.Clear();

            foreach (MainDataGridEntry dgridentry in gridDataEntries)
            {
                dgridentry.Status = "Searching...";
            }
            // Refresh the datagrid to show the new status
            dataGrid.Items.Refresh();

            dataGrid.IsEnabled = false;

            // Disable the search button so no second search can be triggered during first search is still ongoring
            buttonSearchAD.IsEnabled = false;

            try
            {
                foreach (MainDataGridEntry dgridentry in gridDataEntries)
                {
                    // Clear the TailoredADObjects list before executing the "Search" method to make sure no duplicates are added to the list
                    dgridentry.ADObject.TailoredADObjects.Clear();

                    // Search the AD using just the name (to be DELETED)
                    //dgridentry.ADObject.Search(dgridentry.LookupName, checkBoxSearchGlobalCatalog.IsChecked);

                    // Search the AD using the query string NON async
                    // dgridentry.ADObject.Search(dgridentry.LdapQueryString, checkBoxSearchGlobalCatalog.IsChecked);

                    // Sleep for 200 ms to prevent to many LDAP queries at the same time
                    await Task.Delay(200);
                    // Initiate the AD Search async (to be checked if this is thread safe...)
                    await Task.Factory.StartNew<object>(dgridentry.ADObject.Search(dgridentry.LdapQueryString, checkBoxSearchGlobalCatalog.IsChecked));

                    // Show the error message only if there is one                    
                    if (dgridentry.ADObject.QueryErrorMessage != "")
                    {
                        // Todo: Add this error message to the datagrid status
                        MessageBox.Show("aDO error message from catch: " + dgridentry.ADObject.QueryErrorMessage);
                    }

                    dgridentry.ObjectsFound = dgridentry.ADObject.TailoredADObjects.Count;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            foreach (MainDataGridEntry dgridentry in gridDataEntries)
            {
                dgridentry.Status = "Completed";
            }


            // Pupulate "gridDataFoundEntries" with each single search result and show it in the dataGrid
            foreach (MainDataGridEntry dgridentry in gridDataEntries)
            {
                foreach (List<ADObjectNameValuePair> curADObj in dgridentry.ADObject.TailoredADObjects)
                {
                    MainDataGridEntry tempData = new MainDataGridEntry
                    {
                        LookupName = dgridentry.LookupName,
                        LdapQueryString = dgridentry.LdapQueryString,
                        DistinguishedName = curADObj.First(x => x.PropertyName == "distinguishedname").PropertyValue,
                        ObjectsFound = dgridentry.ADObject.TailoredADObjects.Count,
                        Status = "Search completed"
                    };                                      
                    tempData.ADObject.TailoredADObjects.Add(curADObj);
                    gridDataFoundEntries.Add(tempData);
                }

                // If there was no AD Object found, create an entry for the dataGrid which indicates nothing was found
                if(dgridentry.ADObject.TailoredADObjects.Count == 0)
                {
                    MainDataGridEntry tempData = new MainDataGridEntry
                    {
                        LookupName = dgridentry.LookupName,
                        LdapQueryString = dgridentry.LdapQueryString,
                        DistinguishedName = "No search results",
                        ObjectsFound = dgridentry.ADObject.TailoredADObjects.Count,
                        Status = "Search completed"
                    };                    
                    gridDataFoundEntries.Add(tempData);
                }

                // Clear the NameValuePair list in the current datagrid entry / ADObject since it is no more required. The "TailoredADObject" have been transfered
                // to the "gridDataFoundEntries" variable above.
                dgridentry.ADObject.TailoredADObjects.Clear();
                dgridentry.Status = "-";
            }
            dataGrid.ItemsSource = gridDataFoundEntries;
            dataGrid.Columns.First(x => x.Header.ToString() == "ADObject").Visibility = Visibility.Hidden;
            dataGrid.Columns.First(x => x.Header.ToString() == "LdapQueryString").Visibility = Visibility.Hidden;

            // Refresh the datagrid to show the new status
            // Refresh the datagrid to show how many objects has been found
            dataGrid.Items.Refresh();

            dataGrid.IsEnabled = true;

            // Enable the search button 
            buttonSearchAD.IsEnabled = true;
        }


        // This is (seems to be) required to remove the "ADObject" column after the main window has been loaded
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {            
            // Remove the "ADObject" column: https://www.c-sharpcorner.com/uploadfile/dpatra/hide-un-hide-datagrid-columns-in-wpf/            
            try
            { 
                dataGrid.Columns.First(x => x.Header.ToString() == "ADObject").Visibility = Visibility.Hidden;
            }
            catch (Exception exce)
            {
                //MessageBox.Show(exce.Message);
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Close all the AD Object Windows the user opened before closing the AD Searcher Window
            ADObjectViewWindow.CloseAllWindows();
        }
    }
}
