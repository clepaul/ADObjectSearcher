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
using System.Windows;

namespace ADObjectSearcher
{
    public class ADSearcherLib
    {
        // The DLLEntry method must receive two arguments as shown below
        public void DLLEntry(string[] ClipboardText, string dllCfgFilePath)
        {
            XMLRecord xrec = new XMLRecord(dllCfgFilePath);

            ADObjectSearcherWindow MainDLLWindow = new ADObjectSearcherWindow(ClipboardText, xrec)
            {
                // Set display parameters of the new DLL window
                //MainDLLWindow.Left = 0;
                //MainDLLWindow.Top = 0;
                //MainDLLWindow.Topmost = true;
                Owner = Application.Current.MainWindow,
                ShowInTaskbar = false
            };


            // NOTE: Do not call any GUI stuff here, it will break the loading of the DLL 
            //MessageBox.Show("dllCfgFilePath: " + dllCfgFilePath);

            // Show the window of the DLL as a dialog instead of a regular window
            //MainDLLWindow.Show();
            MainDLLWindow.ShowDialog();
        }
    }
}
