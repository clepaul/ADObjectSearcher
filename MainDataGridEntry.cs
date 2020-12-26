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
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADObjectSearcher
{
    class MainDataGridEntry
    {
        public string LookupName { get; set; }
        public string DistinguishedName { get; set; } = "-";
        public string LdapQueryString { get; set; }
        public string Status { get; set; } = "-";
        public int ObjectsFound { get; set; } = 0;
        //public List<SearchADO> ADOList { get; set; } = new List<SearchADO>();
        public SearchADO ADObject { get; private set; } = new SearchADO();


        public MainDataGridEntry()
        { /*
            if(!noADObjects)
            {
                ADObject = new SearchADO();
            } */
        }
    }
}
