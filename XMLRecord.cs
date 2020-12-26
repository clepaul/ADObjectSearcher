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
using System.Xml.Linq;

namespace ADObjectSearcher
{
    public class XMLRecord
    {
        private string pingCallDelayMs;
        public int PingCallDelay { get; private set; }
        public string CustomMessage { get; private set; }

        public XMLRecord(string xmlConfigPath)
        {
            GetXmlData(xmlConfigPath);
        }


        // Read the XML file
        // Source: http://stackoverflow.com/questions/5604330/xml-parsing-read-a-simple-xml-file-and-retrieve-values
        private void GetXmlData(string xmlConfigPath)
        {
            XDocument doc;
            try
            {
                doc = XDocument.Load(xmlConfigPath);
            }
            catch (Exception e)
            {
                // Todo: check how to handle a exeption while creating an instance               

                // Set the delay to 150ms if there is no valid data in the XML
                PingCallDelay = 150;

                CustomMessage = "";

                return;
            }


            try
            {
                pingCallDelayMs = doc.Root.Element("pingCallDelay") != null ? doc.Root.Element("pingCallDelay").Value : "150";
                CustomMessage = doc.Root.Element("customMessage") != null ? doc.Root.Element("customMessage").Value : "";
            }
            catch
            {
                // Todo: check how to handle a exeption while creating an instance               

                // Set the delay to 150ms if there is no valid data in the XML
                PingCallDelay = 150;

                CustomMessage = "";
                return;
            }


            try
            {
                PingCallDelay = int.Parse(pingCallDelayMs);
            }
            catch (Exception)
            {
                // Set the delay to 150ms if there is no valid data in the XML
                PingCallDelay = 150;
            }

        }
    }
}
