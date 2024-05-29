using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

using VNC;

namespace SupportTools_Visio.Data
{
    /// <summary>
    /// ConfigData
    /// 
    /// This class holds all configuration data read from the configuration information passed
    /// in the Reload() method when the initial UserControl is loaded.
    /// 
    /// To use this class, 
    /// 1. Add code in the Reload() method of the initial UserControl to store the ConfigData
    ///    using the RawXML property.  This is done automatically in ucBase.
    /// 2. Add properties to surface the configuration data.  Mark as static.
    ///    Initialize backing store variables!
    /// </summary>
    /// <remarks>Place only static items in this class.  This Class cannot be instantiated.</remarks>
    /// 
    /// <summary>
    /// This class loads information from the app.config/web.config file
    /// and presents it as properties off the Config class.
    /// There should not be any access to AppSettings in other classes.
    /// All configuration information should be provided by this class.
    /// All validation of settings and throwing of exceptions should be done in this class.
    /// The Set() methods get called if an external program, e.g. MCR_EAC needs to set the
    /// values from another source, so keep them!
    /// </summary>
    ///
    public static class Config
    {
        private static int CLASS_BASE_ERRORNUMBER = ErrorNumbers.CONFIG;

        // TODO: Update this with the root element from the configuration data.
        private static string _SmartsDBConnection = ConfigurationManager.AppSettings["ConnectionString"];
        private static string _RootElementName = "AppConfigData";

        #region Initialization

        public static string RawXML
        {
            get;
            set;
        }

        #endregion

        #region Public Configuration Data

        // TODO: Add static Properties to expose configuration information here

        // Please keep these in alphabetical order
        //
        // Please keep the Property names in sync with the setting string.
        // Spaces get replaced with underscores.  Case matters.
        // e.g. "AD Path" --> AD_Path and _AD_Path.

        private static string _AD_Domain = null;
        public static string AD_Domain
        {
            get
            {
                if (_AD_Domain == null)
                {
                    _AD_Domain = ValidateExists_String("AD_Domain");
                }

                return _AD_Domain;
            }
            set
            {
                _AD_Domain = value;
            }
        }

        private static string _AD_Path = null;
        public static string AD_Path
        {
            get
            {
                if (_AD_Path == null)
                {
                    _AD_Path = ValidateExists_String("AD Path");
                }

                return _AD_Path;
            }
            set
            {
                _AD_Path = value;
            }
        }


        private static string _AD_Server = null;
        public static string AD_Server
        {
            get
            {
                if (_AD_Server == null)
                {
                    _AD_Server = ValidateExists_String("AD Server");
                }

                return _AD_Server;
            }
            set
            {
                _AD_Server = value;
            }
        }

        private static Nullable<Boolean> _AD_Users_AllowAll = null;
        public static Boolean AD_Users_AllowAll
        {
            get
            {
                if (_AD_Users_AllowAll == null)
                {
                    _AD_Users_AllowAll = ValidateExists_Boolean("AD_Users_AllowAll");
                }

                return (bool)_AD_Users_AllowAll;
            }
            set
            {
                _AD_Users_AllowAll = value;
            }
        }

        private static string _ADGroup_Administrators = null;
        public static string ADGroup_Administrators
        {
            get
            {
                if (_ADGroup_Administrators == null)
                {
                    _ADGroup_Administrators = ValidateExists_String("ADGroup_Administrators");
                }

                return _ADGroup_Administrators;
            }
            set
            {
                _ADGroup_Administrators = value;
            }
        }

        private static string _ADGroup_BetaUsers = null;
        public static string ADGroup_BetaUsers
        {
            get
            {
                if (_ADGroup_BetaUsers == null)
                {
                    _ADGroup_BetaUsers = ValidateExists_String("ADGroup_BetaUsers");
                }

                return _ADGroup_BetaUsers;
            }
            set
            {
                _ADGroup_BetaUsers = value;
            }
        }

        private static string _ADGroup_Users = null;
        public static string ADGroup_Users
        {
            get
            {
                if (_ADGroup_Users == null)
                {
                    _ADGroup_Users = ValidateExists_String("ADGroup_Users");
                }

                return _ADGroup_Users;
            }
            set
            {
                _ADGroup_Users = value;
            }
        }

        private static Nullable<Boolean> _ADBypass = null;
        public static Boolean ADBypass
        {
            get
            {
                if (_ADBypass == null)
                {
                    //_ADBypass = ValidateExists_Boolean("ADBypass");

                    // Special case this so we don't throw an exception.
                    // No need to broadcast this capability.

                    string configValue = ConfigurationManager.AppSettings["ADBypass"];

                    if (configValue == null)
                    {
                        return false;
                        //throw new ApplicationException(string.Format("Missing Config File Information: {0}", configString));
                    }

                    _ADBypass = Convert.ToBoolean(configValue);
                }

                return (bool)_ADBypass;
            }
            set
            {
                _ADBypass = value;
            }
        }

        private static string _Admin_Tools_Redirect = null;
        public static string Admin_Tools_Redirect
        {
            get
            {
                if (_Admin_Tools_Redirect == null)
                {
                    _Admin_Tools_Redirect = ValidateExists_String("Admin Tools Redirect");
                }

                return _Admin_Tools_Redirect;
            }
            set
            {
                _Admin_Tools_Redirect = value;
            }
        }

        public static string SmartsDBConnection
        {
            get { return _SmartsDBConnection; }
            set
            {
                _SmartsDBConnection = value;
            }
        }

        // Please keep these in alphabetical order
        //
        // Please keep the Property names in sync with the setting string.
        // Spaces get replaced with underscores.  Case matters.
        // e.g. "AD Path" --> AD_Path and _AD_Path.

        private static string _DefaultUITheme = null;
        public static string DefaultUITheme
        {
            get
            {
                if (_DefaultUITheme == null)
                {
                    _DefaultUITheme = ValidateExists_String("DefaultUITheme");
                }

                return _DefaultUITheme;
            }
            set
            {
                _DefaultUITheme = value;

                SaveConfigSetting(value.ToString(), "DefaultUITheme");
            }
        }

        private static Nullable<Int32> _DefaultUserMode = null;
        public static Int32 DefaultUserMode
        {
            get
            {
                if (_DefaultUserMode == null)
                {
                    _DefaultUserMode = ValidateExists_Int32("DefaultUserMode");
                }

                return (Int32)_DefaultUserMode;
            }
            set
            {
                _DefaultUserMode = value;

                SaveConfigSetting(value.ToString(), "DefaultUserMode");
            }
        }

        private static Nullable<Int32> _DefaultViewMode = null;
        public static Int32 DefaultViewMode
        {
            get
            {
                if (_DefaultViewMode == null)
                {
                    _DefaultViewMode = ValidateExists_Int32("DefaultViewMode");
                }

                return (Int32)_DefaultViewMode;
            }
            set
            {
                _DefaultViewMode = value;

                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                AppSettingsSection app = config.AppSettings;
                app.Settings.Remove("DefaultViewMode");
                app.Settings.Add("DefaultViewMode", value.ToString());
                config.Save(ConfigurationSaveMode.Modified);
            }
        }

        private static string _Email_CC = null;
        public static string Email_CC
        {
            get
            {
                if (_Email_CC == null)
                {
                    _Email_CC = ValidateExists_String("Email CC");
                }

                return _Email_CC;
            }
            set
            {
                _Email_CC = value;
            }
        }

        private static string _Email_From = null;
        public static string Email_From
        {
            get
            {
                if (_Email_From == null)
                {
                    _Email_From = ValidateExists_String("Email From");
                }

                return _Email_From;
            }
            set
            {
                _Email_From = value;
            }
        }

        private static string _Email_To = null;
        public static string Email_To
        {
            get
            {
                if (_Email_To == null)
                {
                    _Email_To = ValidateExists_String("Email To");
                }

                return _Email_To;
            }
            set
            {
                _Email_To = value;
            }
        }

        private static string _Enforce_Security = null;
        public static string Enforce_Security
        {
            get
            {
                if (_Enforce_Security == null)
                {
                    _Enforce_Security = ValidateExists_String("Enforce Security");
                }

                return _Enforce_Security;
            }
            set
            {
                _Enforce_Security = value;
            }
        }

        private static string _RowDetailMode = null;
        public static string RowDetailMode
        {
            get
            {
                if (_RowDetailMode == null)
                {
                    _RowDetailMode = ValidateExists_String("DefaultRowDetailMode");
                }

                return _RowDetailMode;
            }
            set
            {
                _RowDetailMode = value;

                SaveConfigSetting(value, "DefaultRowDetailMode");
            }
        }

        private static string _SMTP_Server = null;
        public static string SMTP_Server
        {
            get
            {
                if (_SMTP_Server == null)
                {
                    _SMTP_Server = ValidateExists_String("SMTP Server");
                }

                return _SMTP_Server;
            }
            set
            {
                _SMTP_Server = value;
            }
        }

        #region Template Samples

        // Cat, Dog, NorthWindDBConnectionString, and DataBaseOneConnection
        // are examples and may be deleted.

        private static string _Cat = "";
        public static string Cat
        {
            get
            {
                return PopulateConfigValue("Cat", ref _Cat);
            }
            set
            {
                _Cat = value;
            }
        }


        private static string _Dog;
        public static string Dog
        {
            get
            {
                return PopulateConfigValue("Dog", ref _Dog);
            }
            set
            {
                _Dog = value;
            }
        }

        private static string _DataBaseOneConnection = "";
        public static string DataBaseOneConnection
        {
            get
            {
                return PopulateConfigValue("DatabaseOneConnection", ref _DataBaseOneConnection);
            }
            set
            {
                _DataBaseOneConnection = value;
            }
        }

        private static string _northWindDBConnectionString = "";
        public static string NorthWindDBConnectionString
        {
            get
            {
                return PopulateConfigValue("NorthWindDBConnection", ref _northWindDBConnectionString);
            }
            set
            {
                _northWindDBConnectionString = value;
            }
        }

        #region Complex Property example

        // This shows how to return a more complex type from the config data.

        // 1. First Define the type

        public class urlinfo
        {
            public string envCode;
            public string url;
            public string description;

            public urlinfo(string EnvCode, string Url, string Description)
            {
                envCode = EnvCode;
                url = Url;
                description = Description;
            }
        }

        // 2. Declare a property to expose the type.

        private static IEnumerable<urlinfo> _URLs;
        public static IEnumerable<urlinfo> URLs
        {
            get
            {
                _URLs = GetUrlList("AppConfigData", "Url");
                return _URLs;
            }
            set
            {
                _URLs = value;
            }
        }

        // 3. Declare a method to extract the type from the config data.

        private static IEnumerable<urlinfo> GetUrlList(string root, string param)
        {
            IEnumerable<urlinfo> urls = null;

            if(RawXML == null)
            {
                MessageBox.Show("Error GetUrlList(). ConfigData has not been loaded in Reload() yet.", "Config Data Error", MessageBoxButtons.OK);
                return urls;
            }

            // Don't think we even need to wrap this in a try catch as it is not evaluated here.

            try
            {
                urls =
                    from item in XDocument.Parse(RawXML).Descendants(root).Descendants(param)
                    select new urlinfo(item.Attribute("EnvCode").Value, item.Value, item.Attribute("Description").Value);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return urls;
        }

        #endregion

        #endregion

        #endregion

        public static string GetAllConfigInfo()
        {
            StringBuilder info = new StringBuilder(4096);

            Type ty = typeof(Config);

            object[] methodArgs = new object[1];

            MethodInfo[] methods = ty.GetMethods();

            foreach(MethodInfo method in methods)
            {
                //Common.WriteToDebugWindow(string.Format("Name: {0} ReturnType: {1}", method.Name, method.ReturnType));

                if(method.Name.StartsWith("get"))
                {
                    string value;

                    switch(method.ReturnType.ToString())
                    {
                        case "System.Boolean":
                            try
                            {
                                value = ((Boolean)method.Invoke(ty, null)).ToString();
                            }
                            catch(Exception)
                            {
                                value = "<NULL>";
                            }
                            break;

                        case "System.String":
                            try
                            {
                                value = (string)method.Invoke(ty, null);
                            }
                            catch(Exception)
                            {
                                value = "<NULL>";
                            }

                            break;

                        case "System.Int32":
                            try
                            {
                                value = ((Int32)method.Invoke(ty, null)).ToString();
                            }
                            catch(Exception)
                            {
                                value = "<NULL>";
                            }
                            break;

                        default:
                            value = "Unknown Return Type:" + method.ReturnType.ToString();
                            break;
                    }

                    // Display get_X() as ReturnType Config.X >Value<
                    info.AppendLine(string.Format("{0,16} {1,50} >{2}<",
                        method.ReturnType, "ConfigData." + method.Name.Substring(4), value));
                }
            }

            return info.ToString();
        }

        #region Internal Methods

        internal static string PopulateConfigValue(string configElement, ref string backingVariable)
        {
            if(null == backingVariable)
            {
                MessageBox.Show(
                    string.Format("{0}.  BackingVariable for {1} not initialized",
                        System.Reflection.MethodInfo.GetCurrentMethod().Name, configElement),
                    "ConfigData Error", MessageBoxButtons.OK);
                return "";
            }

            if(backingVariable.Length == 0)
            {
                backingVariable = GetElementValue(_RootElementName, configElement);

                if(0 == backingVariable.Length)
                {
                    MessageBox.Show(
                        string.Format("{0}.  Could not find {1} in <{2}>.",
                            System.Reflection.MethodInfo.GetCurrentMethod().Name, configElement, _RootElementName),
                        "ConfigData Error", MessageBoxButtons.OK);
                }
            }

            return backingVariable;
        }

        /// <summary>
        /// Looks up an value in an XML <element></element> contained by an ancestor <element></element>
        /// </summary>
        /// <param name="descendantName"></param>
        /// <param name="elementName"></param>
        /// <returns></returns>
        internal static XElement GetDescendantElement(string descendantName, string elementName)
        {
            if(RawXML == null)
            {
                MessageBox.Show("Error GetParameter(). ConfigData has not been loaded in Reload() yet.", "Config Data Error", MessageBoxButtons.OK);
                return null;
            }

            try
            {
                // Using cool LINQ types.
                return XDocument.Parse(RawXML).Descendants(descendantName).Elements(elementName).ElementAt(0);
            }
            catch(Exception)
            {
                // We can land here because the param in the <element>(root) did not exist in the config data.
                return null;
            }
        }

        internal static IEnumerable<XElement> GetDescendantElements(string descendantName, string elementName)
        {
            if(RawXML == null)
            {
                MessageBox.Show("Error GetParameter(). ConfigData has not been loaded in Reload() yet.", "Config Data Error", MessageBoxButtons.OK);
                return null;
            }

            try
            {
                return XDocument.Parse(RawXML).Descendants(descendantName).Elements(elementName);
            }
            catch(Exception)
            {
                // We can land here because the param in the <element>(root) did not exist in the config data.
                return null;
            }
        }

        internal static IEnumerable<XElement> GetDescendants(string descendantName)
        {
            if(RawXML == null)
            {
                MessageBox.Show("Error GetParameter(). ConfigData has not been loaded in Reload() yet.", "Config Data Error", MessageBoxButtons.OK);
                return null;
            }

            try
            {
                return XDocument.Parse(RawXML).Descendants(descendantName);
            }
            catch(Exception)
            {
                // We can land here because the param in the <element>(root) did not exist in the config data.
                return null;
            }
        }

        /// <summary>
        /// Looks up an value in an XML <element></element> contained by an ancestor <element></element>
        /// </summary>
        /// <param name="descendantName"></param>
        /// <param name="elementName"></param>
        /// <returns></returns>
        internal static string GetElementValue(string descendantName, string elementName)
        {
            if(RawXML == null)
            {
                MessageBox.Show("Error GetParameter(). ConfigData has not been loaded in Reload() yet.", "Config Data Error", MessageBoxButtons.OK);
                return "";
            }

            try
            {
                // Using cool LINQ types.
                return XDocument.Parse(RawXML).Descendants(descendantName).Elements(elementName).ElementAt(0).Value;

                // Old school.
                //    xmlDoc.LoadXml(RawXML);
                //    xmlRoot = (XmlElement)xmlDoc.GetElementsByTagName(descendantName).Item(0);
                //    xml = xmlRoot.GetElementsByTagName(elementName).Item(0).InnerXml;
            }
            catch(Exception)
            {
                // We can land here because the param in the <element>(root) did not exist in the config data.
                return "";
            }
        }

        #endregion

        #region Validaters

        private static Boolean ValidateExists_Boolean(string configString)
        {
            string configValue = ConfigurationManager.AppSettings[configString];

            if (configValue == null)
            {
                throw new ApplicationException(string.Format("Missing Config File Information: {0}", configString));
            }

            return Convert.ToBoolean(configValue);
        }

        private static Nullable<DateTime> ValidateExists_DateTime(string configString)
        {
            string configValue = ConfigurationManager.AppSettings[configString];

            if (configValue == null)
            {
                return null;
            }

            try
            {
                Log.Info(string.Format("Using CurrentDate {0} from Config File", configValue), Common.LOG_CATEGORY, CLASS_BASE_ERRORNUMBER + 0);
                return Convert.ToDateTime(configValue);
            }
            catch (Exception ex)
            {
                Log.Error(string.Format("Cannot convert {0} to DateTime {1}", configString, ex.ToString()), Common.LOG_CATEGORY, CLASS_BASE_ERRORNUMBER + 0);
                return null;
            }
        }

        private static int ValidateExists_Int32(string configString)
        {
            string configValue = ConfigurationManager.AppSettings[configString];

            if (configValue == null)
            {
                throw new ApplicationException(string.Format("Missing Config File Information: {0}", configString));
            }

            return Convert.ToInt32(configValue);
        }

        private static string ValidateExists_String(string configString)
        {
            string configValue = ConfigurationManager.AppSettings[configString];

            if (configValue == null)
            {
                throw new ApplicationException(string.Format("Missing Config File Information: {0}", configString));
            }

            // TODO(crhodes): Decide if want to bother reporting that some Config item is empty.
            // Some settings, e.g. Email_CC can be empty.  Don't throw exception or catch it in the
            // particular items that can be empty.  For now just log to get a feel if this is important at all.
            // May just want to remove the next lines completely.

            if (configValue.Length <= 0)
            {
                Log.Info(string.Format("Empty Config File Information: {0}", configString), Common.LOG_CATEGORY, CLASS_BASE_ERRORNUMBER + 5);
                //throw new ApplicationException(string.Format("Empty Config File Information: {0}", configString));
            }

            return configValue;
        }

        #endregion

        #region Helpers

        public static bool IsValid_ConfigInformation()
        {
            bool result = true;

            Type ty = typeof(Config);
            object[] methodArgs = new object[1];

            //PropertyInfo[] properties = ty.GetProperties();

            //foreach(PropertyInfo property in properties)
            //{
            //    System.Diagnostics.Debug.WriteLine(property.Name);
            //    System.Diagnostics.Debug.WriteLine(property.PropertyType.ToString());
            //}

            // Read every property out of the Config class to ensure everything needed is available.
            // It is the responsibility of the properties to validate themselves
            // and throw exceptions if not valid/available.

            MethodInfo[] methods = ty.GetMethods();

            foreach (MethodInfo method in methods)
            {
                //Common.WriteToDebugWindow(string.Format("Name: {0} ReturnType: {1}", method.Name, method.ReturnType));

                if (method.Name.StartsWith("get"))
                {
                    string value = "Unknown Return Type";

                    switch (method.ReturnType.ToString())
                    {
                        case "System.Boolean":
                            try
                            {
                                value = ((Boolean)method.Invoke(ty, null)).ToString();
                            }
                            catch (Exception)
                            {
                                Log.Error(string.Format("Invalid Config File Information: {0}", method.Name), Common.LOG_CATEGORY, CLASS_BASE_ERRORNUMBER + 1);
                                result = false;
                            }

                            break;

                        case "System.String":
                            try
                            {
                                value = (string)method.Invoke(ty, null);
                                if (value == null) value = "<NULL>";
                            }
                            catch (Exception)
                            {
                                Log.Error(string.Format("Invalid Config File Information: {0}", method.Name), Common.LOG_CATEGORY, CLASS_BASE_ERRORNUMBER + 1);
                                result = false;
                            }

                            break;

                        case "System.Int32":
                            try
                            {
                                value = ((Int32)method.Invoke(ty, null)).ToString();
                            }
                            catch (Exception)
                            {
                                Log.Error(string.Format("Invalid Config File Information: {0}", method.Name), Common.LOG_CATEGORY, CLASS_BASE_ERRORNUMBER + 1);
                                result = false;
                            }

                            break;
                    }

                    //Common.WriteToDebugWindow(string.Format("Name:{0}  ReturnType:{1} Value:>{2}<", 
                    //    method.Name, method.ReturnType, value));                   
                }
            }

            return result;
        }

        private static void SaveConfigSetting(string value, string settingName)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            AppSettingsSection app = config.AppSettings;
            app.Settings.Remove(settingName);
            app.Settings.Add(settingName, value.ToString());
            config.Save(ConfigurationSaveMode.Modified);
        }

        #endregion

    }
}

