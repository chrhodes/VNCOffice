using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using VNC;
using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.User_Interface.User_Controls
{
    public partial class wucTaskPane_TFS : UserControl
    {

        XlHlp.XlLocation AddSection_CatalogNode(
            XlHlp.XlLocation insertAt, 
            CatalogNode catalogNode)
        {
            XlHlp.AddLabeledInfo(insertAt.AddRow(), "FullPath:", catalogNode.FullPath);
            XlHlp.AddLabeledInfo(insertAt.AddRow(), "IsDefault:", catalogNode.IsDefault.ToString());

            insertAt.ColumnsAdded = 2;

            if (!insertAt.OrientVertical)
            {
                // Skip past the info just added.
                insertAt.SetLocation(insertAt.RowStart, insertAt.TableEndColumn + 1);
            }

            return insertAt;
        }

        XlHlp.XlLocation AddSection_ChildNodes(
            XlHlp.XlLocation insertAt, 
            string NodeType, 
            ReadOnlyCollection<CatalogNode> childNodes)
        {
            //XlHlp.DisplayInWatchWindow(insertAt);

            //// List the team project collections

            XlHlp.AddContentToCell(insertAt.InsertRow(1), string.Format("{0}({1})", NodeType, childNodes.Count));

            //currentRow = startingRow;
            //int innerRowsAdded = 0;
            //int col = 1;

            //foreach (CatalogNode child in childNodes)
            //{
            //    // Need to fix this so expands down the page
            //    innerRowsAdded = AddSection_CatalogNode(ws, rngOutput, child);
            //    insertAt = innerRowsAdded;
            //    currentRow += innerRowsAdded;
            //}

            //XlLocation.Rows++;

            return insertAt;
        }



        //private XlHlp.XlLocation AddBranches(ItemIdentifier[] items, XlHlp.XlLocation insertAt, int currentColumn)
        //{
        //    XlHlp.DisplayInWatchWindow(insertAt);
        //     XlLocation XlLocation = new XlLocation(0,0);


        //    foreach (var item in items)
        //    {
        //        ExcelHlp.DisplayInWatchWindow(string.Format("  Item.ChangeType:{0}",
        //            item.ChangeType.ToString()));

        //        switch (item.ChangeType)
        //        {
        //            case Microsoft.TeamFoundation.VersionControl.Client.ChangeType.Add:

        //                break;

        //            case Microsoft.TeamFoundation.VersionControl.Client.ChangeType.Branch:

        //                break;

        //            case Microsoft.TeamFoundation.VersionControl.Client.ChangeType.Delete:

        //                break;

        //            case Microsoft.TeamFoundation.VersionControl.Client.ChangeType.Edit:

        //                break;

        //            case Microsoft.TeamFoundation.VersionControl.Client.ChangeType.Encoding:

        //                break;

        //            case Microsoft.TeamFoundation.VersionControl.Client.ChangeType.Lock:

        //                break;

        //            case Microsoft.TeamFoundation.VersionControl.Client.ChangeType.Merge:

        //                break;

        //            case Microsoft.TeamFoundation.VersionControl.Client.ChangeType.None:

        //                break;

        //            case Microsoft.TeamFoundation.VersionControl.Client.ChangeType.Property:

        //                break;

        //            case Microsoft.TeamFoundation.VersionControl.Client.ChangeType.Rename:

        //                break;

        //            case Microsoft.TeamFoundation.VersionControl.Client.ChangeType.Rollback:

        //                break;

        //            case Microsoft.TeamFoundation.VersionControl.Client.ChangeType.SourceRename:

        //                break;

        //            case Microsoft.TeamFoundation.VersionControl.Client.ChangeType.Undelete:

        //                break;
        //        }
        //    }

        //XlHlp.DisplayInWatchWindow(System.Reflection.MethodInfo.GetCurrentMethod().Name, insertAt, "End");
        //    return insertAt;
        //}

        XlHlp.XlLocation AddSection_TeamManager_Info(
            XlHlp.XlLocation insertAt)
        {
            long startTicks = XlHlp.DisplayInWatchWindow(insertAt);

            try
            {

                // TODO(crhodes)
                // Add Dummy output
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            XlHlp.DisplayInWatchWindow(insertAt, startTicks, "End");

            return insertAt;
        }
    }
}
