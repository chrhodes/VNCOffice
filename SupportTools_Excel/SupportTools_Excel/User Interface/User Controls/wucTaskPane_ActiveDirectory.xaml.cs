using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using DevExpress.Xpf.Editors;

namespace SupportTools_Excel.User_Interface.User_Controls
{
    /// <summary>
    /// Interaction logic for wucTastkPane_ActiveDirectory.xaml
    /// </summary>
    public partial class wucTaskPane_ActiveDirectory : UserControl
    {
        #region Fields and Properties


        #endregion

        #region Constructors and Load

        public wucTaskPane_ActiveDirectory()
        {
            InitializeComponent();
            LoadControlContents();
        }

        private void LoadControlContents()
        {
            try
            {
                wucActiveDirectory_Picker.PopulateControlFromFile(Common.cCONFIG_FILE);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            //wucSQLInstance_Picker1.ControlChanged += WucSQLInstance_Picker1_ControlChanged;
            //wucTFSProvider_Picker.ControlChanged += tfsProvider_Picker1_ControlChanged;
        }

        private void WucSQLInstance_Picker1_ControlChanged()
        {
            //VNC.AddinHelper.Common.WriteToDebugWindow("wucSQLInstance_Picker1.ControlChanged");
        }

        #endregion

        #region Event Handlers

        private void btnAddUser_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnCheckUserExists_Click(object sender, RoutedEventArgs e)
        {
            string userName = teSearchPattern.Text;

            teOutput.Text = VNC.ActiveDirectory.Helper.UserExists(userName).ToString();
        }

        private void btnFindName_Click(object sender, RoutedEventArgs e)
        {
            teOutput.Text = VNC.ActiveDirectory.Helper.FindName(teSearchPattern.Text);
        }

        private void btnFindUser_Click(object sender, RoutedEventArgs e)
        {
            string userName = teSearchPattern.Text;

            try
            {
                // create LDAP connection object  

                var name = wucActiveDirectory_Picker.Name;
                var dnsHostName = wucActiveDirectory_Picker.DNSHostName;
                var defaultNamingContext = wucActiveDirectory_Picker.DefaultNamingContext;

                DirectoryEntry myLdapConnection = CreateDirectoryEntry(dnsHostName, defaultNamingContext);

                // create search object which operates on LDAP connection object  
                // and set search object to only find the user specified  

                DirectorySearcher search = new DirectorySearcher(myLdapConnection);
                search.Filter = "(cn=" + userName + ")";

                //List<String> requiredProperties = new List<string>();


                SearchResult result = search.FindOne();

                if (result != null)
                {
                    lbeResults.Clear();

                    // user exists, cycle through LDAP fields (cn, telephonenumber etc.)  

                    ResultPropertyCollection fields = result.Properties;

                    foreach (String ldapField in fields.PropertyNames)
                    {
                        // cycle through objects in each field e.g. group membership  
                        // (for many fields there will only be one object such as name)  

                        foreach (Object myCollection in fields[ldapField])
                        {
                            string field = String.Format(">{0,30}< : >{1}<\n", ldapField.Trim(), myCollection.ToString());
                            lbeResults.Text += field;
                        }
                    }
                }

                else
                {
                    Console.WriteLine("User not found!");
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine("Exception caught:\n\n" + ex.ToString());
            }
        }

        UserPrincipal SearchBy(UserPrincipal userPrincipal, string searchBy, string searchPattern)
        {
            switch (searchBy)
            {
                case "Name":
                    userPrincipal.Name = searchPattern;
                    break;

                case "SAMAccountName":
                    userPrincipal.SamAccountName = searchPattern;
                    break;

                default:
                    break;
                   
            }

            return userPrincipal;
        }
        private void btnFindUser2_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                lbeResults.Clear();

                var name = wucActiveDirectory_Picker.Name;
                var dnsHostName = wucActiveDirectory_Picker.DNSHostName;
                var defaultNamingContext = wucActiveDirectory_Picker.DefaultNamingContext;

                // enter AD settings  
                using (PrincipalContext AD = new PrincipalContext(ContextType.Domain, dnsHostName))
                {
                    UserPrincipal userPrincipal = new UserPrincipal(AD);
                    UserPrincipal result = default;

                    userPrincipal = SearchBy(userPrincipal, ((ListBoxEditItem)lbeSearchBy.SelectedItem).Content.ToString(), teSearchPattern.Text); ; ;

                    PrincipalSearcher search = new PrincipalSearcher(userPrincipal);
                    userPrincipal.Dispose();

                    switch (((ListBoxEditItem)lbeFindCount.SelectedItem).Content.ToString())
                    {
                        case "FindOne":
                            result = (UserPrincipal)search.FindOne();
                            search.Dispose();
                            DisplayResults(result);
                            break;

                        case "FindAll":
                            PrincipalSearchResult<Principal> results = search.FindAll();
                            search.Dispose();
                            // TODO(crhodes)
                            // Display Results
                            break;
                    }
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        private void DisplayResults(UserPrincipal result)
        {
            lbeResults.Text += string.Format("{0,30} : >{1}<\n", "Description", result.Description);
            lbeResults.Text += string.Format("{0,30} : >{1}<\n", "Surname", result.Surname);
            lbeResults.Text += string.Format("{0,30} : >{1}<\n", "HomeDirectory", result.HomeDirectory);
            lbeResults.Text += string.Format("{0,30} : >{1}<\n", "HomeDirectory", result.HomeDrive);
            lbeResults.Text += string.Format("{0,30} : >{1}<\n", "DisplayName", result.DisplayName);
            lbeResults.Text += string.Format("{0,30} : >{1}<\n", "MiddleName", result.MiddleName);
            lbeResults.Text += string.Format("{0,30} : >{1}<\n", "SamAccountName", result.SamAccountName);
            lbeResults.Text += string.Format("{0,30} : >{1}<\n", "DistinguishedName", result.DistinguishedName);
            lbeResults.Text += string.Format("{0,30} : >{1}<\n", "GivenName", result.GivenName);
            lbeResults.Text += string.Format("{0,30} : >{1}<\n", "Name", result.Name);
            lbeResults.Text += string.Format("{0,30} : >{1}<\n", "Sid", result.Sid);
            lbeResults.Text += string.Format("{0,30} : >{1}<\n", "UserPrincipalName", result.UserPrincipalName);
            lbeResults.Text += string.Format("{0,30} : >{1}<\n", "ScriptPath", result.ScriptPath);

            // Get lower level object so can access additional properties
            DirectoryEntry lowerLdap = (DirectoryEntry)result.GetUnderlyingObject();

            foreach (string property in lowerLdap.Properties.PropertyNames)
            {
                lbeResults.Text += string.Format("{0,30} : >{1}<\n", property, lowerLdap.Properties[property][0].ToString());
            }
        }

        private void btnFindUserProperties_Click(object sender, RoutedEventArgs e)
        {
            string userName = teSearchPattern.Text;

            try
            {
                // create LDAP connection object  

                var name = wucActiveDirectory_Picker.Name;
                var dnsHostName = wucActiveDirectory_Picker.DNSHostName;
                var defaultNamingContext = wucActiveDirectory_Picker.DefaultNamingContext;

                using (DirectoryEntry myLdapConnection = CreateDirectoryEntry(dnsHostName, defaultNamingContext))
                {
                    // create search object which operates on LDAP connection object  
                    // and set search object to only find the user specified  

                    using (DirectorySearcher search = new DirectorySearcher(myLdapConnection))
                    {
                        search.Filter = "(cn=" + userName + ")";

                        List<String> requiredProperties = new List<string>();

                        //string[] requiredProperties = new string[cbeAttributes.SelectedItems.Count()];

                        foreach (ComboBoxEditItem item in cbeAttributes.SelectedItems)
                        {
                            requiredProperties.Add(item.Content.ToString());
                            search.PropertiesToLoad.Add(item.Content.ToString());
                        }

                        // create results objects from search object  

                        SearchResult result = search.FindOne();

                        if (result != null)
                        {
                            lbeResults.Clear();

                            foreach (string property in requiredProperties)
                            {
                                foreach (Object myCollection in result.Properties[property])
                                {
                                    string value = String.Format(">{0,30}< : >{1}<\n", property, myCollection.ToString());
                                    lbeResults.Text += value;
                                    //lbeResults.Items.Add(field);
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("User not found!");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught:\n\n" + ex.ToString());
            }
        }

        private void btnGetAllGroups_Click(object sender, RoutedEventArgs e)
        {
            System.Collections.ArrayList al = VNC.ActiveDirectory.Helper.GetAllADDomainGroups();

            StringBuilder sb = new StringBuilder();

            foreach (var item in al)
            {
                sb.AppendLine(item.ToString());
            }

            MessageBox.Show(sb.ToString());
        }

        static DirectoryEntry CreateDirectoryEntry(string dnsHostName, string defaultNamingContext)
        {
            // create and return new LDAP connection with desired settings  

            DirectoryEntry ldapConnection = new DirectoryEntry(dnsHostName);
            //DirectoryEntry ldapConnection = new DirectoryEntry();
            //ldapConnection.Path = "LDAP://OU=BDUsers,OU=AMERICAS,OU=CA154,DC=bdx,DC=com";
            ldapConnection.Path = string.Format("LDAP://{0}", defaultNamingContext);
            //ldapConnection.Path = @"LDAP://CN=Users,DC=VNC,DC=LOCAL";
            ldapConnection.AuthenticationType = AuthenticationTypes.Secure;

            return ldapConnection;
        }

        private void btnGetAllUsers_Click(object sender, RoutedEventArgs e)
        {
            System.Collections.ArrayList al = VNC.ActiveDirectory.Helper.GetAllADDomainUsers();

            StringBuilder sb = new StringBuilder();

            foreach (var item in al)
            {
                sb.AppendLine(item.ToString());
            }

            MessageBox.Show(sb.ToString());
        }

        private void btnGetAllUsersPath_Click(object sender, RoutedEventArgs e)
        {
            string domainPath = tePath.Text;

            System.Collections.ArrayList al = VNC.ActiveDirectory.Helper.GetAllADDomainUsers(domainPath);

            StringBuilder sb = new StringBuilder();

            foreach (var item in al)
            {
                sb.AppendLine(item.ToString());
            }

            MessageBox.Show(sb.ToString());
        }

        private void btnGetDomainControllers_Click(object sender, RoutedEventArgs e)
        {
            System.Collections.ArrayList al = VNC.ActiveDirectory.Helper.EnumerateDomainControllers();

            StringBuilder sb = new StringBuilder();

            foreach (var item in al)
            {
                sb.AppendLine(item.ToString());
            }

            MessageBox.Show(sb.ToString());
        }

        private void btnGetDomains_Click(object sender, RoutedEventArgs e)
        {
            System.Collections.ArrayList al = VNC.ActiveDirectory.Helper.EnumerateDomains();

            StringBuilder sb = new StringBuilder();

            foreach (var item in al)
            {
                sb.AppendLine(item.ToString());
            }

            MessageBox.Show(sb.ToString());
        }

        private void btnGetGlobalCatalogs_Click(object sender, RoutedEventArgs e)
        {
            System.Collections.ArrayList al = VNC.ActiveDirectory.Helper.EnumerateGlobalCatalogs();

            StringBuilder sb = new StringBuilder();

            foreach (var item in al)
            {
                sb.AppendLine(item.ToString());
            }

            MessageBox.Show(sb.ToString());
        }

        private void btnLogoff_Click(object sender, RoutedEventArgs e)
        {
            Logoff();
            //btnLogon.Enabled = true;
            //btnLogon.BackColor = SystemColors.Control;
            //btnLogoff.Enabled = false;
            //btnLogoff.BackColor = SystemColors.Control;
            //lblInstancName.Text = "";
            //gbInstanceOperations.Visible = false;
        }

        private void btnLogon_Click(object sender, RoutedEventArgs e)
        {
            if (Logon() == true)
            {
                //btnLogoff.Enabled = true;
                //btnLogoff.BackColor = Color.Green;
                //btnLogon.Enabled = false;
                //btnLogon.BackColor = Color.Green;
                //lblInstancName.Text = ucDBInstanceList.InstanceName;
                //gbInstanceOperations.Visible = true;
            }
        }

        #endregion

        #region Main Function Routines

        #region CreateWorksheet_*



        #endregion

        #region AddSection_*


        #endregion

        #region Display_*

    

        #endregion

        private void Logoff()
        {
            try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private bool Logon()
        {
            bool result = false;

            try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return result;
        }

        #endregion

        #region Utility Routines



        #endregion

        #region Private Methods

        private bool GetDisplayOrientation()
        {
            return (bool)ceOrientOutputVertically.IsChecked;
        }

        private bool ValidUISelections()
        {
            //if (cbeTeamProjectCollections.SelectedText.Length > 0)
            //{
            return true;
            //}
            //else
            //{
            //    MessageBox.Show("Must Select Team Project Collection first", "UI Selection Incomplete");
            //    return false;
            //}
        }

        private void lbeXXX_SelectedIndexChanged(object sender, RoutedEventArgs e)
        {

        }


        #endregion

        //private void lbeResults_Scroll(object sender, System.Windows.Controls.Primitives.ScrollEventArgs e)
        //{

        //}
    }
}
