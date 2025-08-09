using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SupportTools_Visio.User_Interface
{
    public class CustomFormat
    {
        const double GB = 1024 * 1024 * 1024;
        const double MB = 1024 * 1024;
        const double KB = 1024;

        public static void FormatStorageColumns(DevExpress.Xpf.Grid.CustomColumnDisplayTextEventArgs e)
        {
            var v = e.Value;
            var t = e.Source.Tag;

            if (e.Value == null) return;
            if (e.Column == null) return;

            double adjustedDoubleValue = 0;

            // Convert values so all display units are MB except for Total_* and Physical Memory which is GB

            try
            {
                switch (e.Column.FieldName)
                {
                    // DataFile - KB
                    // LogFile - KB
                    // Double
                    case "AvailableSpace":
                        adjustedDoubleValue = (double)e.Value / KB;
                        e.DisplayText = string.Format("{0:N1}", adjustedDoubleValue);
                        break;

                    // Double
                    case "DataSpaceUsage":
                    case "DataSpaceUsed":
                        adjustedDoubleValue = (double)e.Value / KB;
                        e.DisplayText = string.Format("{0:N1}", adjustedDoubleValue);
                        break;

                    // DataFile - KB
                    // LogFile - KB
                    // Double
                    case "IndexSpaceUsage":
                        adjustedDoubleValue = (double)e.Value / KB;
                        e.DisplayText = string.Format("{0:N1}", adjustedDoubleValue);
                        break;

                    // Double
                    case "MaxSize":
                        e.DisplayText = string.Format("{0:N1}", e.Value);
                        break;

                    // Int32
                    case "PhysicalMemory":
                        Int32 adjustedInt32Value = (Int32)e.Value / (Int32)KB;
                        e.DisplayText = string.Format("{0}", adjustedInt32Value);
                        break;

                    // Int64
                    case "RowCount":
                        e.DisplayText = string.Format("{0:N0}", e.Value);
                        break;

                    // Database - MB
                    // FileGroup - KB
                    // DataFile - KB
                    // LogFile - KB

                    // Double
                    case "Size":
                        string units = ((e.Source.Tag == null) ? "KB" : "MB");

                        if (units == "KB")
                        {
                            adjustedDoubleValue = (double)e.Value / KB;
                        }
                        else
                        {
                            adjustedDoubleValue = (double)e.Value;
                        }
                        e.DisplayText = string.Format("{0:N1}", adjustedDoubleValue);
                        break;

                    // Double
                    case "SpaceAvailable":
                        adjustedDoubleValue = (double)e.Value / KB;
                        e.DisplayText = string.Format("{0:N1}", adjustedDoubleValue);
                        break;

                    // Double
                    case "Total_DataSpaceUsage":
                        adjustedDoubleValue = (double)e.Value / MB;
                        e.DisplayText = string.Format("{0:N1}", adjustedDoubleValue);
                        break;

                    // DataFile - KB
                    // LogFile - KB
                    // Double
                    case "Total_IndexSpaceUsage":
                        adjustedDoubleValue = (double)e.Value / MB;
                        e.DisplayText = string.Format("{0:N1}", adjustedDoubleValue);
                        break;

                    case "Total_Size":
                        adjustedDoubleValue = (double)e.Value / KB;
                        e.DisplayText = string.Format("{0:N1}", adjustedDoubleValue);
                        break;

                    case "Total_SpaceAvailable":
                        adjustedDoubleValue = (double)e.Value / MB;
                        e.DisplayText = string.Format("{0:N1}", adjustedDoubleValue);
                        break;

                    // DataFile - KB
                    // LogFile = KB
                    // Double
                    case "UsedSpace":
                        adjustedDoubleValue = (double)e.Value / KB;
                        e.DisplayText = string.Format("{0:N1}", adjustedDoubleValue);
                        break;

                    // DataFile - Bytes
                    // LogFile - Bytes
                    // Int64
                    case "VolumeFreeSpace":
                        adjustedDoubleValue = (Int64)e.Value / MB;
                        e.DisplayText = string.Format("{0:N1}", adjustedDoubleValue);
                        break;

                    default:
                        return;
                }
            }
            catch (Exception ex)
            {
                // e not set for some reason.
                e.DisplayText = "??";
            }
        }
    }
}
