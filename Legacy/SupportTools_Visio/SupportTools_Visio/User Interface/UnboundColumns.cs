using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

using DevExpress.Xpf.Grid;

namespace SupportTools_Visio.User_Interface
{
    public class UnboundColumns
    {
        //public static void GetEnvironmentInstanceDatabaseColumns(GridColumnDataEventArgs e)
        //{
        //    if (e.IsGetData)
        //    {
        //        Guid database_ID;
        //        Guid instance_ID;
        //        string instance_ID_String = null;

        //        switch (e.Column.FieldName)
        //        {
        //            case "ADDomain":
        //                if (e.GetListSourceFieldValue("Instance_ID") != null)
        //                {
        //                    instance_ID = Guid.Parse(e.GetListSourceFieldValue("Instance_ID").ToString());
        //                    var instance = Common.ApplicationDataSet.Instances.Where(ins => ins.ID == instance_ID).First();

        //                    e.Value = instance.ADDomain;

        //                }
        //                else if (e.GetListSourceFieldValue("Database_ID") != null)
        //                {
        //                    database_ID = Guid.Parse(e.GetListSourceFieldValue("Database_ID").ToString());
        //                    var database1 = Common.ApplicationDataSet.Databases.Where(db => db.ID == database_ID).First();
        //                    var instance1 = Common.ApplicationDataSet.Instances.Where(ins => ins.ID == database1.Instance_ID).First();

        //                    e.Value = instance1.ADDomain;
        //                }
        //                else
        //                {
        //                    // TODO(crhodes): Who is sending us here>
        //                }

        //                break;

        //            case "Environment":
        //                if (e.GetListSourceFieldValue("Instance_ID") != null)
        //                {
        //                    instance_ID = Guid.Parse(e.GetListSourceFieldValue("Instance_ID").ToString());
        //                    var instance = Common.ApplicationDataSet.Instances.Where(ins => ins.ID == instance_ID).First();

        //                    e.Value = instance.Environment;

        //                }
        //                else if (e.GetListSourceFieldValue("Database_ID") != null)
        //                {
        //                    database_ID = Guid.Parse(e.GetListSourceFieldValue("Database_ID").ToString());
        //                    var database1 = Common.ApplicationDataSet.Databases.Where(db => db.ID == database_ID).First();
        //                    var instance1 = Common.ApplicationDataSet.Instances.Where(ins => ins.ID == database1.Instance_ID).First();

        //                    e.Value = instance1.Environment;
        //                }
        //                else
        //                {
        //                    // TODO(crhodes): Who is sending us here>
        //                }

        //                break;

        //            case "IsMonitored_Instance":
        //                if (e.GetListSourceFieldValue("Instance_ID") != null)
        //                {
        //                    instance_ID = Guid.Parse(e.GetListSourceFieldValue("Instance_ID").ToString());
        //                    var instance = Common.ApplicationDataSet.Instances.Where(ins => ins.ID == instance_ID).First();

        //                    e.Value = instance.IsIsMonitoredNull() ? false : instance.IsMonitored;

        //                }
        //                else
        //                {
        //                    // TODO(crhodes): Who is sending us here>
        //                }

        //                break;

        //            case "Name_Database":
        //                if (e.GetListSourceFieldValue("Database_ID") != null)
        //                {
        //                    database_ID = Guid.Parse(e.GetListSourceFieldValue("Database_ID").ToString());
        //                    var database = Common.ApplicationDataSet.Databases.Where(db => db.ID == database_ID).First();

        //                    e.Value = database.Name_Database;
        //                }
        //                else
        //                {
        //                    // TODO(crhodes): Who is sending us here>
        //                }

        //                break;

        //            case "Name_Instance":
        //                if (e.GetListSourceFieldValue("Database_ID") != null)
        //                {
        //                    database_ID = Guid.Parse(e.GetListSourceFieldValue("Database_ID").ToString());
        //                    var database2 = Common.ApplicationDataSet.Databases.Where(db => db.ID == database_ID).First();
        //                    var instance2 = Common.ApplicationDataSet.Instances.Where(ins => ins.ID == database2.Instance_ID).First();

        //                    e.Value = instance2.Name_Instance;
        //                }
        //                else
        //                {
        //                    // TODO(crhodes): Who is sending us here>
        //                }

        //                break;

        //            case "Notes_Instance":
        //                if (e.GetListSourceFieldValue("Instance_ID") != null)
        //                {
        //                    instance_ID = Guid.Parse(e.GetListSourceFieldValue("Instance_ID").ToString());
        //                    var instance = Common.ApplicationDataSet.Instances.Where(ins => ins.ID == instance_ID).First();

        //                    e.Value = instance.IsNotesNull() ? "" : instance.Notes;

        //                }
        //                else
        //                {
        //                    // TODO(crhodes): Who is sending us here>
        //                }

        //                break;

        //            case "SecurityZone":
        //                if (e.GetListSourceFieldValue("Instance_ID") != null)
        //                {
        //                    instance_ID = Guid.Parse(e.GetListSourceFieldValue("Instance_ID").ToString());
        //                    var instance = Common.ApplicationDataSet.Instances.Where(ins => ins.ID == instance_ID).First();

        //                    e.Value = instance.IsSecurityZoneNull() ? "" : instance.SecurityZone;
        //                    //e.Value = instance.SecurityZone;

        //                }
        //                else if (e.GetListSourceFieldValue("Database_ID") != null)
        //                {
        //                    database_ID = Guid.Parse(e.GetListSourceFieldValue("Database_ID").ToString());
        //                    var database1 = Common.ApplicationDataSet.Databases.Where(db => db.ID == database_ID).First();
        //                    var instance1 = Common.ApplicationDataSet.Instances.Where(ins => ins.ID == database1.Instance_ID).First();

        //                    e.Value = instance1.IsSecurityZoneNull() ? "" : instance1.SecurityZone;
        //                    //e.Value = instance1.SecurityZone;
        //                }
        //                else
        //                {
        //                    // TODO(crhodes): Who is sending us here>
        //                }

        //                break;

        //            case "SnapShotErrorIndicator":
        //                if (e.GetListSourceFieldValue("Instance_ID") != null)
        //                {
        //                    instance_ID = Guid.Parse(e.GetListSourceFieldValue("Instance_ID").ToString());
        //                    var instance = Common.ApplicationDataSet.Instances.Where(ins => ins.ID == instance_ID).First();

        //                    string errorMessage = instance.SnapShotError;

        //                    if (errorMessage.Length > 0)
        //                    {
        //                        if (errorMessage.Equals("Connection Failure Exception"))
        //                        {
        //                            e.Value = GetImage("OrangeBall.png");
        //                        }
        //                        else
        //                        {
        //                            e.Value = GetImage("RedBall.png");
        //                        }
        //                    }
        //                    else
        //                    {
        //                        e.Value = GetImage("GreenBall.png");
        //                    }

        //                }
        //                else
        //                {
        //                    // TODO(crhodes): Who is sending us here>
        //                }

        //                break;

        //            case "Total_DB_Data":
        //                if (e.GetListSourceFieldValue("ID") != null)
        //                {
        //                    instance_ID = Guid.Parse(e.GetListSourceFieldValue("ID").ToString());
        //                    var instance = Common.ApplicationDataSet.Instances.Where(ins => ins.ID == instance_ID).First();

        //                    var calculation = Common.ApplicationDataSet.Databases
        //                        .Where(inst => inst.Instance_ID == instance_ID)
        //                        .Select(inst => inst.DataSpaceUsage)
        //                        .Sum();

        //                    e.Value = calculation.ToString();
        //                }
        //                else
        //                {
        //                    // TODO(crhodes): Who is sending us here>
        //                }  
                      
        //                break;

        //            case "Total_DB_Index":
        //                if (e.GetListSourceFieldValue("ID") != null)
        //                {
        //                    instance_ID = Guid.Parse(e.GetListSourceFieldValue("ID").ToString());
        //                    var instance = Common.ApplicationDataSet.Instances.Where(ins => ins.ID == instance_ID).First();

        //                    var calculation = Common.ApplicationDataSet.Databases
        //                        .Where(inst => inst.Instance_ID == instance_ID)
        //                        .Select(inst => inst.IndexSpaceUsage)
        //                        .Sum();

        //                    e.Value = calculation.ToString();
        //                }
        //                else
        //                {
        //                    // TODO(crhodes): Who is sending us here>
        //                }

        //                break;

        //            case "Total_DB_Size":
        //                if (e.GetListSourceFieldValue("ID") != null)
        //                {
        //                    instance_ID = Guid.Parse(e.GetListSourceFieldValue("ID").ToString());
        //                    var instance = Common.ApplicationDataSet.Instances.Where(ins => ins.ID == instance_ID).First();

        //                    var calculation = Common.ApplicationDataSet.Databases
        //                        .Where(inst => inst.Instance_ID == instance_ID)
        //                        .Select(inst => inst.Size)
        //                        .Sum();

        //                    e.Value = calculation.ToString();
        //                }
        //                else
        //                {
        //                    // TODO(crhodes): Who is sending us here>
        //                }

        //                break;

        //            case "Total_DB_SpaceAvailable":
        //                if (e.GetListSourceFieldValue("ID") != null)
        //                {
        //                    instance_ID = Guid.Parse(e.GetListSourceFieldValue("ID").ToString());
        //                    var instance = Common.ApplicationDataSet.Instances.Where(ins => ins.ID == instance_ID).First();

        //                    var calculation = Common.ApplicationDataSet.Databases
        //                        .Where(inst => inst.Instance_ID == instance_ID)
        //                        .Select(inst => inst.SpaceAvailable)
        //                        .Sum();

        //                    e.Value = calculation.ToString();
        //                }
        //                else
        //                {
        //                    // TODO(crhodes): Who is sending us here>
        //                }

        //                break;
        //        }
        //    }
        //}

        public static void GetSnapShotErrorColumns(GridColumnDataEventArgs e)
        {
            if (e.IsGetData)
            {
                
                switch (e.Column.FieldName)
                {
                    case "SnapShotErrorIndicator":
                        string errorMessage = e.GetListSourceFieldValue("SnapShotError").ToString();

                        if (errorMessage.Length > 0)
                        {
                            if (errorMessage.Equals("Connection Failure Exception"))
                            {
                                e.Value = GetImage("OrangeBall.png");
                            }
                            else
                            {
                                e.Value = GetImage("RedBall.png");                                
                            }
                        }
                        else
                        {
                            e.Value = GetImage("GreenBall.png");
                        }
                        
                        break;
                }
            }
        }

        static BitmapFrame GetImage(string resourceName)
        {
            Uri uri = new Uri("pack://application:,,,/Resources/Images/" + resourceName, UriKind.Absolute);
            return BitmapFrame.Create(uri);
        }

    }
}
