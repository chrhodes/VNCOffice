using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

using AH = VNC.AssemblyHelper;
using XlHlp = VNC.AddinHelper.Excel;

using Crc32C;
using System.IO;

namespace SupportTools_Excel.User_Interface.User_Controls
{
    /// <summary>
    /// Interaction logic for wucTaskPane_Developer.xaml
    /// </summary>
    public partial class wucTaskPane_Developer : UserControl
    {
        #region Fields and Properties

        AH.AssemblyInformation _assemblyInfo;
        Assembly _assembly = null;
        AH.AssemblyReflectionManager _assemblyReflectionManager = new AH.AssemblyReflectionManager();
        #endregion

        #region Constructors and Load

        public wucTaskPane_Developer()
        {
            InitializeComponent();
            LoadControlContents();
        }

        private void LoadControlContents()
        {
            try
            {
                PopulateControlFromFile(Common.cCONFIG_FILE);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            //wucAssembly_Picker1.ControlChanged += wucAssembly_Picker1_ControlChanged;
            //wucTFSProvider_Picker.ControlChanged += tfsProvider_Picker1_ControlChanged;
        }

        private void Reload_Click(object sender, RoutedEventArgs e)
        {
            LoadDataFromFile();
        }

        private void LoadDataFromFile()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
            openFileDialog1.FileName = "";

            if (true == openFileDialog1.ShowDialog())
            {
                string fileName = openFileDialog1.FileName;

                PopulateControlFromFile(fileName);
            }
        }

        public void PopulateControlFromFile(string fileNameAndPath)
        {
            assemblyList.Source = null;
            cbeAssemblies.Items.Clear();
            assemblyList.Source = new Uri(fileNameAndPath);
        }

        #endregion

        #region Event Handlers
        private void cbeAssemblies_PopupClosed(object sender, DevExpress.Xpf.Editors.ClosePopupEventArgs e)
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));
        }
        private void cbeAssemblies_PopupOpened(object sender, RoutedEventArgs e)
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));
        }
        private void cbeAssemblies_EditorActivated(object sender, RoutedEventArgs e)
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));
        }
        private void cbeAssemblies_EditValueChanged(object sender, DevExpress.Xpf.Editors.EditValueChangedEventArgs e)
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            PopulateSelectedTypes(cbeAssemblies);
        }
        private void cbeAssemblies_SelectedIndexChanged(object sender, RoutedEventArgs e)
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));
        }
        private void cbeAssemblies_SourceUpdated(object sender, System.Windows.Data.DataTransferEventArgs e)
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));
        }
        private void cbeAssemblies_DefaultButtonClick(object sender, RoutedEventArgs e)
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));
        }

        

        private void btnCreateWS_Master_Assembly_TypeInfo_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidUISelections()) { return; }

            bool orientVertical = GetDisplayOrientation();

            try
            {
                XlHlp.ScreenUpdatesOff();

                CreateWS_Master_Assembly_TypeInfo(cbeAssemblies);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                XlHlp.ScreenUpdatesOn(true);
            }
        }

        private void btnCreateWS_Master_Assembly_TypeMethodInfo_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidUISelections()) { return; }

            bool orientVertical = GetDisplayOrientation();

            try
            {
                XlHlp.ScreenUpdatesOff();

                CreateWS_Master_Assembly_TypeMethodInfo(cbeAssemblies, (bool)ceAllTypes.IsChecked);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                XlHlp.ScreenUpdatesOn(true);
            }
        }

        void CreateWS_Master_Assembly_ValueTypeInfo(DevExpress.Xpf.Editors.ComboBoxEdit comboBoxEdit, bool allTypes)
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}",
                 System.Reflection.MethodInfo.GetCurrentMethod().Name));

            string sheetName;

            if (cbeAssemblies.SelectedItems.Count > 1)
            {
                sheetName = XlHlp.SafeSheetName(string.Format("{0}{1}-{2}", "MAVTI>", "What", "Else"));
            }
            else
            {
                System.Xml.XmlElement item = (System.Xml.XmlElement)cbeAssemblies.SelectedItem;

                sheetName = XlHlp.SafeSheetName(string.Format("{0}{1}-{2}", "MAVTI>", item.Attributes["SourceName"].Value, item.Attributes["ApplicationName"].Value));
            }

            Worksheet ws = XlHlp.NewWorksheet(sheetName, afterSheetName: "LAST");

            XlHlp.XlLocation insertAt = new XlHlp.XlLocation(ws, row: 5, column: 1, orientVertical: true);

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            AddHeaderFor_Assembly_ValueTypeInfo(insertAt);

            insertAt.IncrementRows();

            foreach (System.Xml.XmlElement item in cbeAssemblies.SelectedItems)
            {
                string assemblyPath = string.Format("{0}\\{1}", item.Attributes["RootPath"].Value, item.Attributes["FileName"].Value);
                string sourceName = item.Attributes["SourceName"].Value;
                string applicationName = item.Attributes["ApplicationName"].Value;
                string reflectionDomain = string.Format("{0}_{1}_{2}", "VNCReflectionDomain", sourceName, applicationName);

                try
                {
                    LoadAssemblyFromPath(assemblyPath, reflectionDomain);

                    foreach (AH.ValueTypeInformation structure in _assemblyReflectionManager.GetStructureInformation(assemblyPath))
                    {
                        insertAt.ClearOffsets();

                        if (sourceName != "")
                        {
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), sourceName);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), applicationName);
                        }

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), structure.Assembly);
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), structure.DeclaringType);
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), structure.Type);

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), structure.Field);
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), structure.FieldType);
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), structure.IsArray);
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), structure.IsEnum);
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), structure.IsValueType);

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), structure.Attribute);
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), structure.AttributeValue);
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), structure.AttributeToString);

                        insertAt.IncrementRows();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), assemblyPath);
                }
            }

            insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblMasterInfoAll_{0}", "ValueType"));

            insertAt.Group(insertAt.OrientVertical, hide: true);

            insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
        }

        private void btnCreateWS_Master_Assembly_ValueTypeInfo_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidUISelections()) { return; }

            bool orientVertical = GetDisplayOrientation();

            try
            {
                XlHlp.ScreenUpdatesOff();

                CreateWS_Master_Assembly_ValueTypeInfo(cbeAssemblies, (bool)ceAllTypes.IsChecked);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                XlHlp.ScreenUpdatesOn(true);
            }
        }

        struct AssemblyDetails
        {
            public string SourceName;
            public string ApplicationName;
            public AssemblyDetails(string sourceName, string applicationName)
            {
                this.SourceName = sourceName;
                this.ApplicationName = applicationName;
            }
        }

        #endregion

        #region Main Function Routines

        #region CreateWS_*
      
        private void CreateWS_Master_Assembly_TypeInfo(DevExpress.Xpf.Editors.ComboBoxEdit cbeAssemblies)
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            string sheetName;
            string sourceName = "";
            string applicationName = "";

            if (cbeAssemblies.SelectedItems.Count > 1)
            {
                sheetName = XlHlp.SafeSheetName(string.Format("{0}{1}-{2}", "MATI>", "What", "Else"));
            }
            else
            {
                System.Xml.XmlElement item = (System.Xml.XmlElement)cbeAssemblies.SelectedItem;



                sheetName = XlHlp.SafeSheetName(string.Format("{0}{1}-{2}", "MATI>", sourceName, applicationName));
            }

            Worksheet ws = XlHlp.NewWorksheet(sheetName, afterSheetName: "LAST");

            XlHlp.XlLocation insertAt = new XlHlp.XlLocation(ws, row: 5, column: 1, orientVertical: true);

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            AddHeaderFor_Assembly_TypeInfo(insertAt);

            insertAt.IncrementRows();
            
            foreach (System.Xml.XmlElement item in cbeAssemblies.SelectedItems)
            {
                string assemblyPath = string.Format("{0}\\{1}", item.Attributes["RootPath"].Value, item.Attributes["FileName"].Value);

                sourceName = item.Attributes["SourceName"].Value;
                applicationName = item.Attributes["ApplicationName"].Value;

                string reflectionDomain = string.Format("{0}_{1}_{2}", "VNCReflectionDomain", sourceName, applicationName);

                try
                {
                    if (LoadAssemblyFromPath(assemblyPath, reflectionDomain))
                    {
                        string source = item.Attributes["SourceName"].Value;
                        string application = item.Attributes["ApplicationName"].Value;
                        string assemblyFileName = item.Attributes["FileName"].Value;

                        insertAt = DisplayListOf_DefinedTypes2(insertAt, assemblyPath, assemblyFileName, source, application);
                        //insertAt = DisplayListOf_DefinedTypes(insertAt, assembly, source, application);
                    }

                    //assembly = null;
                    // TODO(crhodes)
                    // We could call
                    //_assemblyReflectionManager.UnloadDomain(reflectionDomain);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), assemblyPath);
                }
            }

            insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblMasterInfoAll_{0}", "Type"));

            insertAt.Group(insertAt.OrientVertical, hide: true);

            insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
        }

        private void CreateWS_Master_Assembly_TypeMethodInfo(DevExpress.Xpf.Editors.ComboBoxEdit cbeAssemblies, bool allTYpes)
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            string sheetName;

            if (cbeAssemblies.SelectedItems.Count > 1)
            {
                sheetName = XlHlp.SafeSheetName(string.Format("{0}{1}-{2}", "MATMI>", "What", "Else"));
            }
            else
            {
                System.Xml.XmlElement item = (System.Xml.XmlElement)cbeAssemblies.SelectedItem;

                sheetName = XlHlp.SafeSheetName(string.Format("{0}{1}-{2}", "MATMI>", item.Attributes["SourceName"].Value, item.Attributes["ApplicationName"].Value));
            }

            Worksheet ws = XlHlp.NewWorksheet(sheetName, afterSheetName: "LAST");

            XlHlp.XlLocation insertAt = new XlHlp.XlLocation(ws, row: 5, column: 1, orientVertical: true);

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            AddHeaderFor_Assembly_TypeMethodInfo(insertAt);

            insertAt.IncrementRows();

            foreach (System.Xml.XmlElement item in cbeAssemblies.SelectedItems)
            {
                string assemblyPath = string.Format("{0}\\{1}", item.Attributes["RootPath"].Value, item.Attributes["FileName"].Value);
                string sourceName = item.Attributes["SourceName"].Value;
                string applicationName = item.Attributes["ApplicationName"].Value;
                string reflectionDomain = string.Format("{0}_{1}_{2}", "VNCReflectionDomain", sourceName, applicationName);

                XlHlp.DisplayInWatchWindow(String.Format("LoadFrom({0})", assemblyPath));

                try
                {
                    LoadAssemblyFromPath(assemblyPath, reflectionDomain);

                    using (MD5 md5Hash = MD5.Create())
                    {

                        foreach (AH.MethodInformation method in _assemblyReflectionManager.GetMethodInformation(assemblyPath))
                        {
                            insertAt.ClearOffsets();

                            if (sourceName != "")
                            {
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), sourceName);
                                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), applicationName);
                            }

                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), method.Assembly);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), method.Type);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), method.IsStatic);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), method.IsPublic);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), method.IsPrivate);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), method.ReturnType);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), method.Method);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), method.Parameters);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), method.RetValParameters);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), method.OutParameters);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), method.OptionalParameters);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), method.ByRefParameters);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), method.MD5);

                            insertAt.IncrementRows();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), assemblyPath);
                }
            }

            insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblMasterInfoAll_{0}", "TypeMethod"));

            insertAt.Group(insertAt.OrientVertical, hide: true);

            insertAt.EndSectionAndSetNextLocation(insertAt.OrientVertical);
        }

        string GetMd5Hash(MD5 md5Hash, byte[] bytes)
        {
            // Convert the input string to a byte array and compute the hash.
            byte[] data = md5Hash.ComputeHash(bytes);

            StringBuilder sBuilder = new StringBuilder();

            // Loop through each byte of the hashed data 
            // and format each one as a hexadecimal string.

            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }

            return sBuilder.ToString();
        }

        private static string FormatMethodParameters(MethodInfo method, 
            ref bool hasRetValParameters, ref bool hasOutParameters, 
            ref bool hasOptionalParameters, ref bool hasByRefParameters)
        {
            string output = "";
            bool isFirstParameter = true;

            if (method.GetParameters().Count() > 0)
            {
                foreach (ParameterInfo parameter in method.GetParameters())
                {
                    output += string.Format("{0}{1} {2}", 
                        isFirstParameter ? "" : ", ",
                        parameter.ParameterType, parameter.Name);

                    isFirstParameter = false;

                    if (parameter.IsRetval) { hasRetValParameters = true; }
                    if (parameter.IsOut) { hasOutParameters = true; }
                    if (parameter.IsOptional) { hasOptionalParameters = true; }
                    if (parameter.ParameterType.IsByRef) { hasByRefParameters = true; }
                }
            }
            else
            {
                output = "()";
            }

            return output;
        }

        #endregion

        #region AddHeaderFor_*

        private static void AddHeaderFor_Assembly_TypeInfo(XlHlp.XlLocation insertAt)
        {
            string[] headers = AH.TypeInformation.GetHeadersSourceContext();

            foreach (string item in headers)
            {
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item);
            }
        }

        private static void AddHeaderFor_Assembly_TypeMethodInfo(XlHlp.XlLocation insertAt)
        {
            string[] headers = AH.MethodInformation.GetHeadersSourceContext();

            foreach (string item in headers)
            {
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item);
            }
        }

        private static void AddHeaderFor_Assembly_ValueTypeInfo(XlHlp.XlLocation insertAt)
        {
            string[] headers = AH.ValueTypeInformation.GetHeadersSourceContext();

            foreach (string item in headers)
            {
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item);
            }
        }

        //private static void AddHeaderFor_DefinedTypes_Info(XlHlp.XlLocation insertAt)
        //{
        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "Assembly");

        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "Type");
        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "DeclaringType");
        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "Name");

        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "IsPublic");
        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "IsNotPublic");

        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "IsValueType");

        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "IsPrimitive");
        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "IsEnum");
        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "IsInterface");

        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "IsClass");
        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "IsAbstract");

        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "IsSealed");
        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "IsNested");
        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "IsNestedPublic");
        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "IsNestedPrivate");

        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "HasElementType");
        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "IsArray");
        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "IsByRef");
        //    XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), "IsPointer");
        //}

        private static void AddHeaderFor_Type_Info(XlHlp.XlLocation insertAt)
        {
            string[] headers = AH.TypeInformation.GetHeaders();

            foreach (string item in headers)
            {
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), item);
            }
        }

        #endregion

        #region AddSection_*

        private XlHlp.XlLocation AddSection_CustomAttributes_Info(XlHlp.XlLocation insertAt, Assembly assembly)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);
            Worksheet ws = insertAt.workSheet;

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "Custom Attributes");

            insertAt.IncrementRows();

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            insertAt = DisplayListOf_CustomAttributes(insertAt, assembly.CustomAttributes);

            insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblCustomAttributes_{0}", ws.Name));

            insertAt.Group(insertAt.OrientVertical, hide: true);

            if (!insertAt.OrientVertical)
            {
                // Skip past the info just added.
                insertAt.SetLocation(insertAt.RowStart, insertAt.TableEndColumn + 1);
            }

            return insertAt;
        }

        private XlHlp.XlLocation AddSection_DefinedTypes_Info(XlHlp.XlLocation insertAt, Assembly assembly)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);
            Worksheet ws = insertAt.workSheet;

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "DefinedTypes");

            insertAt.IncrementRows();

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            //AddHeaderFor_DefinedTypes_Info(insertAt);
            AddHeaderFor_Type_Info(insertAt);

            insertAt.IncrementRows();

            try
            {
                //insertAt = DisplayListOf_DefinedTypes(insertAt, assembly, "", "");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblDefinedTypes_{0}", ws.Name));

            insertAt.Group(insertAt.OrientVertical, hide: true);

            if (!insertAt.OrientVertical)
            {
                // Skip past the info just added.
                insertAt.SetLocation(insertAt.RowStart, insertAt.TableEndColumn + 1);
            }

            return insertAt;
        }

        private XlHlp.XlLocation AddSection_ExportedTypes_Info(XlHlp.XlLocation insertAt, Assembly assembly)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);
            Worksheet ws = insertAt.workSheet;

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "ExportedTypes");

            insertAt.IncrementRows();

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            AddHeaderFor_Type_Info(insertAt);

            insertAt.IncrementRows();

            try
            {
                insertAt = DisplayListOf_ExportedTypes(insertAt, assembly);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblExportedTypes_{0}", ws.Name));

            insertAt.Group(insertAt.OrientVertical, hide: true);

            if (!insertAt.OrientVertical)
            {
                // Skip past the info just added.
                insertAt.SetLocation(insertAt.RowStart, insertAt.TableEndColumn + 1);
            }

            return insertAt;
        }

        private XlHlp.XlLocation AddSection_Method_Info(XlHlp.XlLocation insertAt, Assembly assembly, TypeInfo typeInfo, string sourceName, string applicationName, MD5 md5Hash)
        {
            try
            {
                Type type = assembly.GetType(typeInfo.FullName);

                BindingFlags bindingFlags =
                    BindingFlags.Public
                    | BindingFlags.NonPublic
                    | BindingFlags.DeclaredOnly
                    | BindingFlags.Instance
                    | BindingFlags.Static;

                foreach (MethodInfo methodInfo in type.GetMethods(bindingFlags))
                {
                    try
                    {
                        insertAt.ClearOffsets();
                        string methodParameters = "";

                        bool hasRetValParameters = false;
                        bool hasOutParameters = false;
                        bool hasOptionalParameters = false;
                        bool hasByRefParameters = false;

                        methodParameters = FormatMethodParameters(methodInfo,
                            ref hasRetValParameters, ref hasOutParameters, ref hasOptionalParameters, ref hasByRefParameters);

                        if (sourceName != "")
                        {
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), sourceName);
                            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), applicationName);
                        }

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), assembly.GetName().Name);

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.FullName);

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), methodInfo.IsStatic.ToString());
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), methodInfo.IsPublic.ToString());
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), methodInfo.IsPrivate.ToString());

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), methodInfo.ReturnType.ToString());

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), methodInfo.Name);

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), methodParameters);

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), hasRetValParameters ? "X" : "");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), hasOutParameters ? "X" : "");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), hasOptionalParameters ? "X" : "");
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), hasByRefParameters ? "X" : "");

                        byte[] methodBodyBytes = methodInfo.GetMethodBody().GetILAsByteArray();

                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), Crc32CAlgorithm.Compute(methodBodyBytes).ToString());
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), GetMd5Hash(md5Hash, methodBodyBytes));                     
                    }
                    catch (Exception ex)
                    {
                        XlHlp.DisplayInWatchWindow(ex.ToString());
                    }
                    finally
                    {
                        insertAt.IncrementRows();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), typeInfo.FullName);
            }

            return insertAt;
        }

        private XlHlp.XlLocation AddSection_Module_Info(XlHlp.XlLocation insertAt, Assembly assembly)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);
            Worksheet ws = insertAt.workSheet;

            //XlHlp.AddTitledInfo(insertAt.AddRow(), "Module Information for ", assembly.FullName);

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "Modules");

            insertAt.IncrementRows();

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            insertAt = DisplayListOf_Modules(insertAt, assembly.Modules);

            insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblModules_{0}", ws.Name));

            insertAt.Group(insertAt.OrientVertical, hide: true);

            if (!insertAt.OrientVertical)
            {
                // Skip past the info just added.
                insertAt.SetLocation(insertAt.RowStart, insertAt.TableEndColumn + 1);
            }

            return insertAt;
        }

        private XlHlp.XlLocation AddSection_References_Info(XlHlp.XlLocation insertAt, Assembly assembly)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);
            Worksheet ws = insertAt.workSheet;

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "References");

            insertAt.IncrementRows();

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            insertAt = DisplayListOf_References(insertAt, assembly.GetReferencedAssemblies());

            insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblModules_{0}", ws.Name));

            insertAt.Group(insertAt.OrientVertical, hide: true);

            if (!insertAt.OrientVertical)
            {
                // Skip past the info just added.
                insertAt.SetLocation(insertAt.RowStart, insertAt.TableEndColumn + 1);
            }

            return insertAt;
        }

        private XlHlp.XlLocation AddSection_Types_Info(XlHlp.XlLocation insertAt, Assembly assembly)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            Worksheet ws = insertAt.workSheet;

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "Types");

            insertAt.IncrementRows();

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            //insertAt = DisplayListOf_Types(insertAt, assembly.GetTypes());

            insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblTypes_{0}", ws.Name));

            insertAt.Group(insertAt.OrientVertical, hide: true);

            if (!insertAt.OrientVertical)
            {
                // Skip past the info just added.
                insertAt.SetLocation(insertAt.RowStart, insertAt.TableEndColumn + 1);
            }

            return insertAt;
        }
        
        private XlHlp.XlLocation AddSection_X(XlHlp.XlLocation insertAt, Collection<string> strings)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            XlHlp.AddContentToCell(insertAt.AddRowX(), "Team Projects");

            Worksheet ws = insertAt.workSheet;

            XlHlp.AddColumnHeaderToSheetX(insertAt.AddOffsetColumnX(), columnWidth: 25, "Name");

            insertAt.IncrementRows();

            insertAt.MarkStart(XlHlp.MarkType.GroupTable);

            insertAt = DisplayListOf_X(insertAt, strings);

            insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tblTP_{0}", ws.Name));

            insertAt.Group(insertAt.OrientVertical, hide: true);

            if (!insertAt.OrientVertical)
            {
                // Skip past the info just added.
                insertAt.SetLocation(insertAt.RowStart, insertAt.TableEndColumn + 1);
            }

            return insertAt;
        }

        #endregion

        #region Display_*

        private XlHlp.XlLocation DisplayListOf_CustomAttributes(XlHlp.XlLocation insertAt, IEnumerable<CustomAttributeData> customAttributes)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            // The columns in this method need to be kept in sync with AddSection_Modules()

            foreach (CustomAttributeData customAttributeData in customAttributes)
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), customAttributeData.ToString());

                insertAt.IncrementRows();
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        //private XlHlp.XlLocation DisplayListOf_DefinedTypes(XlHlp.XlLocation insertAt, Assembly assembly, string sourceName = "", string applicationName = "")
        //{
        //    XlHlp.DisplayInWatchWindow(insertAt);

        //    // The columns in this method need to be kept in sync with AddSection_Modules()

            
        //    foreach (TypeInfo typeInfo in assembly.DefinedTypes)
        //    {
        //        insertAt.ClearOffsets();

        //        if (sourceName != "")
        //        {
        //            // Only the Master sheets display source and application
        //            XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), sourceName);
        //            XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), applicationName);
        //        }

        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.Assembly.GetName().Name);

        //        Type type = typeInfo;
        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), type.FullName);
        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.DeclaringType != null ? typeInfo.DeclaringType.Name : "");
        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.Name);

        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsPublic.ToString());
        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsNotPublic.ToString());

        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsValueType.ToString());

        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsPrimitive.ToString());
        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsEnum.ToString());
        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsInterface.ToString());

        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsClass.ToString());
        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsAbstract.ToString());

        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsSealed.ToString());
        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsNested.ToString());
        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsNestedPublic.ToString());
        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsNestedPrivate.ToString());

        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.HasElementType.ToString());
        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsArray.ToString());
        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsByRef.ToString());
        //        XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsPointer.ToString());

        //        insertAt.IncrementRows();
        //    }

        //    XlHlp.DisplayInWatchWindow(System.Reflection.MethodInfo.GetCurrentMethod().Name, insertAt, "End");

        //    return insertAt;
        //}

        private XlHlp.XlLocation DisplayListOf_DefinedTypes2(XlHlp.XlLocation insertAt, string assemblyPath, string assemblyFileName, string sourceName = "", string applicationName = "")
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            // The columns in this method need to be kept in sync with AddSection_Modules()

            foreach (AH.TypeInformation typeInfo in _assemblyReflectionManager.GetTypeInformation(assemblyPath))
            {
                insertAt.ClearOffsets();

                if (sourceName != "")
                {
                    // Only the Master sheets display source and application
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), sourceName);
                    XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), applicationName);
                }

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), assemblyFileName);
                //XlHlp.AddContentToCell(insertAt.AddOffsetColumn(), typeInfo.Assembly.GetName().Name);

                //Type type = typeInfo;
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.FullName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.DeclaringType != null ? typeInfo.DeclaringType : "");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.Name);

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsPublic);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsNotPublic);

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsValueType);

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsPrimitive);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsEnum);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsInterface);

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsClass);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsAbstract);

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsSealed);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsNested);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsNestedPublic);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsNestedPrivate);

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.HasElementType);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsArray);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsByRef);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.IsPointer);

                insertAt.IncrementRows();
            }

            XlHlp.DisplayInWatchWindow(insertAt, "End");

            return insertAt;
        }

        private XlHlp.XlLocation DisplayListOf_ExportedTypes(XlHlp.XlLocation insertAt, Assembly assembly)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            // The columns in this method need to be kept in sync with AddSection_ExportedTypes()

            foreach (Type type in assembly.ExportedTypes)
            {
                insertAt.ClearOffsets();
      
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.Assembly.GetName().Name);

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.FullName);
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.DeclaringType != null ?  type.DeclaringType.Name : "");
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.Name);

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.IsPublic.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.IsNotPublic.ToString());

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.IsValueType.ToString());

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.IsPrimitive.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.IsEnum.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.IsInterface.ToString());

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.IsClass.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.IsAbstract.ToString());

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.IsSealed.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.IsNested.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.IsNestedPublic.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.IsNestedPrivate.ToString());

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.HasElementType.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.IsArray.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.IsByRef.ToString());
                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), type.IsPointer.ToString());

                insertAt.IncrementRows();
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }
        
        private XlHlp.XlLocation DisplayListOf_Modules(XlHlp.XlLocation insertAt, IEnumerable<System.Reflection.Module> modules)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            // The columns in this method need to be kept in sync with AddSection_Modules()

            foreach (System.Reflection.Module module in modules)
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), module.Name);

                insertAt.IncrementRows();
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        private XlHlp.XlLocation DisplayListOf_References(XlHlp.XlLocation insertAt, AssemblyName[] references)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            // The columns in this method need to be kept in sync with AddSection_Modules()

            foreach (AssemblyName assemblyName in references.OrderBy(x => x.FullName))
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), assemblyName.FullName);

                insertAt.IncrementRows();
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        private XlHlp.XlLocation DisplayListOf_Types(XlHlp.XlLocation insertAt, IEnumerable<TypeInfo> types)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            // The columns in this method need to be kept in sync with AddSection_Modules()

            foreach (TypeInfo typeInfo in types)
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), typeInfo.Name);

                insertAt.IncrementRows();
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        private XlHlp.XlLocation DisplayListOf_X(XlHlp.XlLocation insertAt, Collection<string> strings)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            // The columns in this method need to be kept in sync with CreateTeamProjectsInfo()

            foreach (string s in strings)
            {
                insertAt.ClearOffsets();

                XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), s);

                insertAt.IncrementRows();
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }

        #endregion

        #endregion

        #region Utility Routines
        private Boolean LoadAssemblyFromPath(string assemblyPath, string reflectionDomain = "VNC")
        {
            //Assembly assembly = null;
            Boolean result = false;

            ////_assemblyReflectionManager.LoadAssembly(assemblyName, "VNCReflectionDomain");

            //DevExpress.Xpf.Editors.ListBoxEditItem lbeItem = (DevExpress.Xpf.Editors.ListBoxEditItem)lbeAssemblyLoad.SelectedItem;

            //switch (lbeItem.Content.ToString().ToUpper())
            //{
            //    case "LOADFROM":
                    result = _assemblyReflectionManager.LoadAssembly(assemblyPath, reflectionDomain);


            //        //assembly = Assembly.LoadFrom(assemblyPath);
            //        break;

            //    case "LOADFILE":
            //        assembly = Assembly.LoadFile(assemblyPath);
            //        break;

            //    case "REFLECTIONONLYLOADFROM":
            //        assembly = Assembly.ReflectionOnlyLoadFrom(assemblyPath);
            //        break;

            //    default:
            //        MessageBox.Show(lbeAssemblyLoad.ToString(), "Unexpected lbeLoadAssembly EditItem");
            //        break;

            //}

            return result;
        }


        #endregion

        #region Private Methods

        private void PopulateSelectedTypes(DevExpress.Xpf.Editors.ComboBoxEdit cbeAssemblies)
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            cbeAssemblies.Items.Clear();

            foreach (System.Xml.XmlElement item in cbeAssemblies.SelectedItems)
            {
                string assemblyPath = string.Format("{0}\\{1}", item.Attributes["RootPath"].Value, item.Attributes["FileName"].Value);

                XlHlp.DisplayInWatchWindow(String.Format("LoadFrom({0})", assemblyPath));

                try
                {
                    // TODO(crhodes)
                    // Need to decide if should switch this to _assemblyReflectionManager
                    // For now skip.

                    //Assembly assembly = LoadAssemblyFromPath(assemblyPath);

                    //foreach (TypeInfo typeInfo in assembly.DefinedTypes)
                    //{
                    //    if (! cbeSelectedTypes.Items.Contains(typeInfo.Name))
                    //    {
                    //        cbeSelectedTypes.Items.Add(typeInfo.Name);
                    //    }
                    //}
                }
                catch (ReflectionTypeLoadException rtle)
                {
                    MessageBox.Show(rtle.ToString(), "See DeveloperDebug Window");

                    foreach (var ex in rtle.LoaderExceptions)
                    {
                        XlHlp.DisplayInWatchWindow(string.Format("{0}()", ex.ToString()));
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), assemblyPath);
                }
            }
        }

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


        #endregion

        private void btnCreateWS_Assembly_Info_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidUISelections()) { return; }

            bool orientVertical = GetDisplayOrientation();

            try
            {
                XlHlp.ScreenUpdatesOff();

                foreach (System.Xml.XmlElement item in cbeAssemblies.SelectedItems)
                {
                    string assemblyName = string.Format("{0}\\{1}", item.Attributes["RootPath"].Value, item.Attributes["FileName"].Value);

                    XlHlp.DisplayInWatchWindow(String.Format("LoadFrom({0})", assemblyName));

                    try
                    {
                        XlHlp.DisplayInWatchWindow(item.Attributes["SourceName"].Value);
                        XlHlp.DisplayInWatchWindow(item.Attributes["ApplicationName"].Value);

                        Assembly assembly = Assembly.LoadFrom(assemblyName);

                        CreateWS_AssemblyInfo(orientVertical, assembly,
                            new AssemblyDetails(item.Attributes["SourceName"].Value, item.Attributes["ApplicationName"].Value));
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString(), assemblyName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                XlHlp.ScreenUpdatesOn(true);
            }
        }

        private void CreateWS_AssemblyInfo(bool orientVertical, Assembly assembly, AssemblyDetails assemblyDetails)
        {
            XlHlp.DisplayInWatchWindow(string.Format("{0}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            string assemblyName = assembly.GetName().Name;

            string sheetName = XlHlp.SafeSheetName(string.Format("{0}{1}-{2}", "AI>", assemblyName, assemblyDetails.ApplicationName));
            Worksheet ws = XlHlp.NewWorksheet(sheetName, beforeSheetName: "FIRST");

            XlHlp.XlLocation insertAt = new XlHlp.XlLocation(ws, row: 5, column: 1, orientVertical: orientVertical);

            XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Source", assemblyDetails.SourceName);
            XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "ApplicationName", assemblyDetails.ApplicationName);
            XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Assembly Information for ", assembly.FullName);

            insertAt = AddSection_Module_Info(insertAt, assembly);

            insertAt.IncrementRows();

            insertAt = AddSection_References_Info(insertAt, assembly);

            //insertAt.IncrementRows();

            //insertAt = AddSection_DefinedTypes_Info(insertAt, assembly);

            insertAt.IncrementRows();

            insertAt = AddSection_CustomAttributes_Info(insertAt, assembly);

            //insertAt.IncrementRows();

            //insertAt = AddSection_ExportedTypes_Info(insertAt, assembly);
        }
    }
}
