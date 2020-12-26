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
using System.DirectoryServices;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.DirectoryServices.ActiveDirectory;
using System.Text.RegularExpressions;

namespace ADObjectSearcher
{
    // Source: https://stackoverflow.com/questions/456523/quick-way-to-retrieve-user-information-active-directory
    partial class SearchADO
    {  
        public List<List<ADObjectNameValuePair>> TailoredADObjects { get; private set; } = new List<List<ADObjectNameValuePair>>();
        public bool QuerySuccessfullyCompleted { get; private set; } = true;
        public string QueryErrorMessage { get; private set; } = "";
        

        public SearchADO()
        {
            // to be defined
        }

        public static string GetCurrentDomain()
        {
            DirectoryEntry dEntry = new DirectoryEntry("LDAP://RootDSE");

            /*
            DirectoryEntry newRootDomain = new DirectoryEntry($"GC://{dEntry.Properties["rootDomainNamingContext"].Value}");
            MessageBox.Show("Rootdomain.name: " + newRootDomain.Name);
            MessageBox.Show("Rootdomain.Path: " + newRootDomain.Path);
            MessageBox.Show("Rootdomain.SchemaclassName: " + newRootDomain.SchemaClassName);
            MessageBox.Show(dEntry.Properties["rootDomainNamingContext"].PropertyName + ": " + dEntry.Properties["rootDomainNamingContext"].Value); */
            
            if(dEntry.Properties["rootDomainNamingContext"].Value.ToString() == "" || dEntry.Properties["rootDomainNamingContext"].Value == null)
            {
                return "Failed to get domain name.";
            }
            else
            {
                return dEntry.Properties["rootDomainNamingContext"].Value.ToString();
            }            
        }

        private string ExtractSearchDomain(string distinguishedName)
        {
            distinguishedName = distinguishedName.ToUpper();
            string[] splittedLDAPPath = distinguishedName.Split(',');
            string[] domainPartOnly  = splittedLDAPPath.Where(x => x.Contains("DC=")).ToArray();            
            string ldapDomainName = "LDAP://" + String.Join(",", domainPartOnly);
            return ldapDomainName;
        }


        // Search the AD first using Global Catalog. If selected by the user search directly in the target domain to get more details.
        public Func<object> Search(string ldapSearchString, bool? searchInGCOnly = false)
        {
            SearchResultCollection sResults = null;
            SearchResult searchResultTargetDomain = null;

            try
            {
                // Get the root domain to search the Global Catalog
                DirectoryEntry dEntry = new DirectoryEntry("LDAP://RootDSE");
                DirectoryEntry RootDomain = new DirectoryEntry($"GC://{dEntry.Properties["rootDomainNamingContext"].Value}");
                
                DirectorySearcher dSearcher = new DirectorySearcher(RootDomain);
                
                // Set the ldap filter to the value from the datagrid
                dSearcher.Filter = $"(&(|(objectClass=user)(objectClass=group)(objectClass=computer)){ldapSearchString})";

                // To be checked:
                //dSearcher.PageSize = 1000;
                //dSearcher.SizeLimit = 1000;

                // Search AD using Global Catalog
                sResults = dSearcher.FindAll();

                
                // This is the loop for each found AD object in the Global Catalog
                foreach (SearchResult sRes in sResults)
                {
                    // Only search the GC and return those properties
                    if (searchInGCOnly == true)
                    {                        
                        List<ADObjectNameValuePair> adObjValueProps = new List<ADObjectNameValuePair>();

                        // This is the loop to get each single AD object property and its name
                        foreach (string propName in sRes.Properties.PropertyNames)
                        {
                            // Since each property is a collection this loop is for each of the single entries in that collection
                            foreach (object prop in sRes.Properties[propName])
                            {
                                ADObjectNameValuePair curNameValuePair = new ADObjectNameValuePair
                                {
                                    PropertyName = propName,
                                    PropertyValue = prop.ToString(),                                    
                                };

                                // Populated the "CalculatedValue" var / column 
                                curNameValuePair.CalculatedValue = CreateCalculatedValue(propName, prop);

                                adObjValueProps.Add(curNameValuePair);
                            }
                        }

                        TailoredADObjects.Add(adObjValueProps);
                    }
                    // If user selected to search in the domain instead of the GC, perform the domain specific search below
                    else
                    {
                        string targetSearchDomain = ExtractSearchDomain(sRes.Path);
                        object dnToSearch = sRes.Properties["distinguishedname"][0];

                        // Create a DirectrySearcher for the Domain the object is member of
                        DirectoryEntry dEntryTargetDomain = new DirectoryEntry(targetSearchDomain);
                        DirectorySearcher dSearcherTargetDomain = new DirectorySearcher(dEntryTargetDomain);

                        // Set the ldap filter containing the search using the "distinguishedname" which should be unique
                        // Todo: check if "distinguishedname" is really unique
                        dSearcherTargetDomain.Filter = $"(&(|(objectClass=user)(objectClass=group)(objectClass=computer))((distinguishedname={dnToSearch})))";

                        // Search AD using using the target domain, NOT the Global Catalog
                        searchResultTargetDomain = dSearcherTargetDomain.FindOne();

                        List<ADObjectNameValuePair> adObjValueProps = new List<ADObjectNameValuePair>();

                        // This is the loop to get each single AD object property and its name
                        foreach (string propName in searchResultTargetDomain.Properties.PropertyNames)
                        {                            
                            // Since each property is a collection this loop is for each of the single entries in that collection
                            foreach (object prop in searchResultTargetDomain.Properties[propName])
                            {
                                ADObjectNameValuePair curNameValuePair = new ADObjectNameValuePair
                                {
                                    PropertyName = propName,
                                    PropertyValue = prop.ToString()
                                };

                                // Populated the "CalculatedValue" var / column 
                                curNameValuePair.CalculatedValue = CreateCalculatedValue(propName, prop);

                                adObjValueProps.Add(curNameValuePair);                                
                            }
                        }
                        TailoredADObjects.Add(adObjValueProps);
                    }                    
                }
            }
            catch (InvalidOperationException iOe)
            {
                // MessageBox.Show("in FIRST catch: " + iOe.Message);
                QueryErrorMessage = iOe.Message;
                QuerySuccessfullyCompleted = false;
            }

            catch (NotSupportedException nSe)
            {
                // MessageBox.Show("in SECOND catch: " + nSe.Message);
                QueryErrorMessage = nSe.Message;
                QuerySuccessfullyCompleted = false;
            }

            catch (Exception ex)
            {
                // MessageBox.Show("in LAST catch: " + ex.Message);
                QueryErrorMessage = ex.Message;
                QuerySuccessfullyCompleted = false;
            }

            finally
            {
                if (sResults != null)
                {
                    sResults.Dispose();
                }
            }


            // ... required .... to be checked
            //Func<object> func = new Func<object>(new string() blah);
            //Func<object> retObject = func;
            //return retObject ;

            return (Func<object>)(() =>
            {                
                return new object();
            });

        }


        // Create the "Calculated Values"
        private string CreateCalculatedValue(string propName, object prop)
        {
            string theCalculatedValue = "";
            
            switch (propName)
            {
                case "pwdlastset":
                case "lastlogon":
                case "lastlogontimestamp":
                    {
                        // source: https://stackoverflow.com/questions/18614810/how-to-convert-active-directory-pwdlastset-to-date-time
                        // ("accountexpires" is not a "FileTimeUtc" value)                        
                        theCalculatedValue = DateTime.FromFileTimeUtc((long)prop).ToString();                        
                        break;
                    }
                case "objectguid":
                    {
                        // Source: https://stackoverflow.com/questions/18383843/how-do-i-convert-an-active-directory-objectguid-to-a-readable-string/31040455

                        theCalculatedValue = new Guid((byte[])prop).ToString();                        
                        break;
                    }
                case "objectsid":
                    {
                        // Populated the "_Calculated" properties for the SIDs 
                        // source: https://stackoverflow.com/questions/11580128/how-to-convert-sid-to-string-in-net                       
                        theCalculatedValue = new SecurityIdentifier((byte[])prop, 0).ToString();                        
                        break;
                    }
                case "accountexpires":
                    {
                        // Populated the "_Calculated" properties for the accountexpires 
                        // Source: https://stackoverflow.com/questions/6360284/convert-ldap-accountexpires-to-datetime-in-c-sharp
                        // Source: https://stackoverflow.com/questions/8042398/c-sharp-active-directory-accountexpires-property-not-reading-correctly
                        theCalculatedValue = (long.MaxValue == (long)prop) ? "Never" : new DateTime(1601, 01, 01, 0, 0, 0, DateTimeKind.Utc).AddTicks((long)prop).ToString();                       
                        break;
                    }
                case "useraccountcontrol":
                    {
                        // Source: https://stackoverflow.com/questions/10231914/useraccountcontrol-in-active-directory
                        UserAccountControl userAccountControl = (UserAccountControl)prop;
                        // This gets a comma separated string of the flag names that apply.
                        string userAccountControlFlagNames = userAccountControl.ToString();

                        theCalculatedValue = userAccountControlFlagNames;                       
                        break;
                    }
                case "member":
                case "memberof":
                    {   
                        try
                        {
                            // To be simplified / fixed / debugged                            
                            string[] tmpDNArr = prop.ToString().Split(',');
                            theCalculatedValue = tmpDNArr[0].Substring(3);
                        }
                        catch(Exception curEx)
                        {
                            theCalculatedValue = curEx.Message;
                        }                        
                        break;
                    }
            }
            return theCalculatedValue;
        }
    }
}
