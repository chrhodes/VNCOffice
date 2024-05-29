using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

using Visio = Microsoft.Office.Interop.Visio;

namespace SupportTools_Visio
{
    class Helper
    {
        private static void ProcessX(XElement command)
        {
            XElement thingX = command.Element("ThingX");
            string x;
            string y;
            int z;

            foreach (XElement item in thingX.Elements())
            {
                switch (item.Name.ToString())
                {
                    case "DeleteAll":
                        
                        break;

                    case "Operation1":
                        x = item.Attribute("x").Value;
                        z = int.Parse(item.Attribute("z").Value);
                        break;
                }
            }

        }

        private static void ProcessY(XElement command)
        {
            

        }

        static public void ProcessXMLCommand(XElement command)
        {
            ProcessX(command);
            ProcessY(command);
        }

        public static void IndicateApplicationUsage(string application, DateTime eventDate, string user, string message)
        {
            var dataRow = Common.ApplicationDataSet.ApplicationUsage.NewApplicationUsageRow();

            dataRow.Application = application;
            dataRow.EventDate = eventDate;
            dataRow.User = user;
            dataRow.EventMessage = message;

            Common.ApplicationDataSet.ApplicationUsage.AddApplicationUsageRow(dataRow);
            // HACK(crhodes)
            //+++ Skip writing to database for now
            //Common.ApplicationDataSet.ApplicationUsageTA.Update(Common.ApplicationDataSet.ApplicationUsage);
        }

        //+  TODO(crhodes)
        // Should we have more of these for other data types?

        static public string GetShapePropertyAsString(Visio.Shape activeShape, string property)
        {
            string propertyName = "Prop." + property;
            string result = "";

            if (activeShape.CellExistsU[propertyName, 0] != 0)
            {
                result = activeShape.CellsU[propertyName].ResultStrU[Visio.VisUnitCodes.visUnitsString];
            }

            return result;
        }
    }

    //private void ProcessProvisioningFile()
    //{
    //    Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();

    //    XElement hISSchema;

    //    openFileDialog.Filter = "XML Files (*.xml)|*.xml|All files (*.*)|(*.*)";
    //    openFileDialog.DefaultExt = "xml";

    //    if (openFileDialog.ShowDialog() == true)
    //    {
    //        using (var streamReader = new StreamReader(openFileDialog.FileName))
    //        {
    //            hISSchema = XElement.Load(streamReader);
    //            HIS.Library.Helper.ProcessXML(hISSchema);
    //        }
    //    }
    //}

    //private void LoadSets()
    //{
    //    DebugWindow.Common.WriteToDebugWindow(string.Format("{0}:{1}()", CONTROL_NAME, System.Reflection.MethodInfo.GetCurrentMethod().Name));

    //    foreach (XElement setInfo in _ShowXml.Element("Sets").Elements("Set"))
    //    {
    //        Set set = new Set();

    //        set.SetXml = setInfo;
    //        set.DisplayLights = _DisplayLights;

    //        // Add the set to the collection
    //        Sets.Add(set);

    //        set.Prepare();
    //    }
    //}
}
