using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;

using Microsoft.Office.Interop.Excel;

using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

using SupportTools_Excel.AzureDevOpsExplorer.Application;
using SupportTools_Excel.AzureDevOpsExplorer.Domain;

using VNC;

using VNCTFS = VNC.TFS;
using XlHlp = VNC.AddinHelper.Excel;

namespace SupportTools_Excel.AzureDevOpsExplorer.Application
{
    class Worksheet_Output
    {
        internal static void CreateWS_All_TPC_LastChangeset(Options_AZDO_TFS options,
            VersionControlServer versionControlServer)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            try
            {
                XlHlp.XlLocation insertAt = CreateNewWorksheet(string.Format("{0}_{1}", "All_TPC", "LastChangeset"),
                    options);

                XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Last Changeset All TeamProjects", AzureDevOpsExplorer.Presentation.Views.Server.TfsTeamProjectCollection.Name);

                insertAt.MarkStart(XlHlp.MarkType.GroupTable);

                Header_VersionControlServer.Add_Changesets(insertAt);

                //Body_VersionControlServer.Add_Changesets(insertAt, options, versionControlServer);

                foreach (var teamProjectName in options.TeamProjects)
                {
                    insertAt.ClearOffsets();

                    long loopTicks = Log.Trace($"Processing {teamProjectName}", Common.LOG_CATEGORY);

                    TeamProject teamProject = VNCTFS.Helper.Get_TeamProject(versionControlServer, teamProjectName.Trim());

                    if (teamProject != null)
                    {
                        Globals.ThisAddIn.Application.StatusBar = $"Processing {teamProject.Name}";

                        Body_VersionControlServer.Add_TP_Changesets(insertAt, options,
                            Presentation.Views.Server.VersionControlServer, teamProject);

                        AZDOHelper.ProcessLoopDelay(options);
                    }
                    else
                    {
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), teamProjectName);
                        XlHlp.AddOffsetContentToCell(insertAt.AddOffsetColumn(), "No VCS Project");
                        insertAt.IncrementRows();
                    }

                    Log.Trace($"EndProcessing {teamProjectName}", Common.LOG_CATEGORY, loopTicks);
                }

                insertAt.MarkEnd(XlHlp.MarkType.GroupTable, string.Format("tbl_{0}", insertAt.workSheet.Name));

                insertAt.Group(insertAt.OrientVertical, hide: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }

        internal static XlHlp.XlLocation CreateNewWorksheet(string sheetName,
            Options_AZDO_TFS options, [CallerMemberName] string callerName = "")
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            string safeSheetName = XlHlp.SafeSheetName(sheetName);
            Worksheet ws = XlHlp.NewWorksheet(safeSheetName, beforeSheetName: "FIRST");

            XlHlp.XlLocation insertAt = new XlHlp.XlLocation(ws, options.StartingRow, options.StartingColumn, options.OrientOutputVertically);
            XlHlp.AddLabeledInfoX(insertAt.AddRowX(), "Date Run", DateTime.Now.ToString());

            if (!options.FormatSpecs.IsInitialized)
            {
                options.FormatSpecs.Initialize(insertAt);
            }

            using (System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog())
            {
                string strOutputFile = null;

                try
                {
                    saveFileDialog.FileName = "AzureDevOpsExplorer.xlsx";
                    //saveFileDialog.InitialDirectory = startingFolder;

                    if (System.Windows.Forms.DialogResult.Cancel == saveFileDialog.ShowDialog())
                    {
                        return insertAt;
                    }
                    else
                    {
                        strOutputFile = saveFileDialog.FileName;
                    }

                    if (string.IsNullOrEmpty(strOutputFile))
                    {
                        return insertAt;
                    }

                    Globals.ThisAddIn.Application.ActiveWorkbook.SaveAs(strOutputFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);

            return insertAt;
        }
    }
}
