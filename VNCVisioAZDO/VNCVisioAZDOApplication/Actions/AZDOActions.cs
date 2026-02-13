using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;

using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;

using VNC;
using VNC.Core;

using VNCVisioToolsApplication.Domain;

using MSVisio = Microsoft.Office.Interop.Visio;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Actions
{
    public partial class AZDOActions
    {
        public static VNC.WPF.Presentation.Dx.Views.DxThemedWindowHost addLinkedWorkItemsHost = null;

        internal static async void AddLinkedWorkItems1(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] vs)
        {
            VNCVisioAddIn.Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            // NOTE(crhodes)
            // Can launch a UI here.  Or earlier.

            //DxThemedWindowHost.DisplayUserControlInHost(ref addLinkedWorkItemsHost,
            //    "Edit Shape Control Points Text",
            //    Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
            //    DxThemedWindowHost.ShowWindowMode.Modeless,
            //    new Presentation.Views.EditControlPoints());

            MSVisio.Page activePage = app.ActivePage;
            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
            var version = WorkItemShapeInfo.WorkItemShapeVersion.V1;

            AddLinkedWorkItems(app, activePage, activeShape, "WI 1", version);
        }
        
        internal static async void AddLinkedWorkItems2(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] vs)
        {
            VNCVisioAddIn.Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            // NOTE(crhodes)
            // Can launch a UI here.  Or earlier.

            //DxThemedWindowHost.DisplayUserControlInHost(ref addLinkedWorkItemsHost,
            //    "Edit Shape Control Points Text",
            //    Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
            //    DxThemedWindowHost.ShowWindowMode.Modeless,
            //    new Presentation.Views.EditControlPoints());

            MSVisio.Page activePage = app.ActivePage;
            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
            string targetShapeName = activeShape.CellsU["Prop.WIShapeName"].ResultStrU[MSVisio.VisUnitCodes.visUnitsString];
            var version = WorkItemShapeInfo.WorkItemShapeVersion.V2;

            AddLinkedWorkItems(app, activePage, activeShape, targetShapeName, version);
        }

        private static bool VerifyRequiredStencils(MSVisio.Application app)
        {
            bool result = false;

            result = VNCVisioAddIn.Helpers.LoadStencil(app, "Azure DevOps.vssx");
            result = VNCVisioAddIn.Helpers.LoadStencil(app, "Page Shapes.vssx");

            return result;
        }

        internal static async void AddLinkedWorkItems(MSVisio.Application app, MSVisio.Page page, MSVisio.Shape shape, 
            string shapeName, WorkItemShapeInfo.WorkItemShapeVersion version)
        {
            if (!VerifyRequiredStencils(app))
            {
               
                return;
            }

            WorkItemShapeInfo activeShapeWorkItemInfo = new WorkItemShapeInfo(shape);

            int id;

            if (int.TryParse(activeShapeWorkItemInfo.ID, out id))
            {
            }
            else
            {
                MessageBox.Show($"Cannot parse ({activeShapeWorkItemInfo.ID}) as WorkItemID");
                return;
            }

            int relatedLinkCount;

            if (int.TryParse(activeShapeWorkItemInfo.RelatedLinkCount, out relatedLinkCount))
            {
            }
            else
            {
                MessageBox.Show($"Cannot parse ({activeShapeWorkItemInfo.RelatedLinkCount}) as RelatedLinkCount");
                return;
            }

            var result = await VNC.AZDO.Helper.QueryWorkItemLinks(activeShapeWorkItemInfo.Organization, id, relatedLinkCount);

            if (result.Count > 0)
            {
                Point initialPosition = GetPosition(shape);
                Point insertionPoint = initialPosition;

                string stencilName = "Azure DevOps.vssx";

                MSVisio.Document linkStencil;
                MSVisio.Master linkMaster = null;

                try
                {
                    linkStencil = app.Documents[stencilName];

                    try
                    {
                        linkMaster = linkStencil.Masters[shapeName];
                    }
                    catch (Exception ex)
                    {
                        VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("  Cannot find Master named:>{0}<", shapeName));
                    }
                }
                catch (Exception ex)
                {
                    VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("  Cannot find open Stencil named:>{0}<", stencilName));
                }

                // TODO(crhodes)
                // Figure out how to get size of shape from master.
                // HACK(crhodes)
                // .25 is for Link counts

                double height = version == WorkItemShapeInfo.WorkItemShapeVersion.V1 ? 0.375 : 0.475;

                WorkItemOffsets workItemOffsets = new WorkItemOffsets(initialPosition, height: height, padX: 0.25, padY: 0.05);

                foreach (var linkedWorkItem in result)
                {
                    // NOTE(crhodes)
                    // This includes the current shape.  Do not add it.
                    // May always be first one.  Maybe loop counter
                    if (linkedWorkItem.Id == id)
                    {
                        continue;
                    }

                    VNCVisioAddIn.Common.WriteToDebugWindow($"{linkedWorkItem.Id} {linkedWorkItem.Fields["System.Title"]}");

                    insertionPoint = AZDOPageLayout.CalculateInsertionPointLinkedWorkItems(initialPosition, linkedWorkItem, activeShapeWorkItemInfo, workItemOffsets);

                    AddNewWorkItemShapeToPage(page, linkMaster, linkedWorkItem, insertionPoint, activeShapeWorkItemInfo, version);
                }
            }

            VNCVisioAddIn.Common.WriteToDebugWindow($"{activeShapeWorkItemInfo}");
        }

        internal static async void GetWorkItemInfo1(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] vs)
        {
            VNCVisioAddIn.Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];

            GetWorkItemInfo(activeShape, WorkItemShapeInfo.WorkItemShapeVersion.V1);
        }

        internal static async void GetWorkItemInfo2(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] vs)
        {
            VNCVisioAddIn.Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];

            GetWorkItemInfo(activeShape, WorkItemShapeInfo.WorkItemShapeVersion.V2);
        }

        internal static async void GetWorkItemInfo(MSVisio.Shape shape, WorkItemShapeInfo.WorkItemShapeVersion version)
        {
            WorkItemShapeInfo shapeInfo = new WorkItemShapeInfo(shape);

            int id = 0;

            if (!int.TryParse(shapeInfo.ID, out id))
            {
                MessageBox.Show($"Invalid WorkItem ID: ({shapeInfo.ID})");
                return;
            }

            var result = await VNC.AZDO.Helper.QueryWorkItemInfoById(shapeInfo.Organization, id);

            if (result.Count == 0)
            {
                MessageBox.Show($"Cannot find WorkItem ID: ({shapeInfo.ID})");
                return;
            }

            shapeInfo.InitializeFromWorkItem(result[0]);

            // NOTE(crhodes)
            // Go add the bugs

            int bugs = await VNC.AZDO.Helper.QueryRelatedBugsById(shapeInfo.Organization, int.Parse(shapeInfo.ID));

            shapeInfo.RelatedBugs = bugs.ToString();

            shapeInfo.PopulateShapeDataFromInfo(shape, version);

            VNCVisioAddIn.Common.WriteToDebugWindow($"{shapeInfo}");
        }

        private static async void AddNewWorkItemShapeToPage(MSVisio.Page page, MSVisio.Master linkMaster,
            WorkItem workItem, Point insertionPoint,
            WorkItemShapeInfo relatedShape, 
            WorkItemShapeInfo.WorkItemShapeVersion version)
        {
            try
            {
                MSVisio.Shape newWorkItemShape = page.Drop(linkMaster, insertionPoint.X, insertionPoint.Y);
                WorkItemShapeInfo shapeInfo = new WorkItemShapeInfo(newWorkItemShape);
                shapeInfo.InitializeFromWorkItem(workItem);

                int bugs = await VNC.AZDO.Helper.QueryRelatedBugsById(shapeInfo.Organization, int.Parse(shapeInfo.ID));

                shapeInfo.RelatedBugs = bugs.ToString();

                shapeInfo.PopulateShapeDataFromInfo(newWorkItemShape, version);
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToDebugWindow($"{workItem.Id} - {ex}");
            }
        }

        private static async void AddNewWorkItemRevisionShapeToPage(MSVisio.Page page, MSVisio.Master linkMaster,
            WorkItem workItem, Point insertionPoint,
            WorkItemShapeInfo relatedShape,
            WorkItemShapeInfo.WorkItemShapeVersion version)
        {
            try
            {
                MSVisio.Shape newWorkItemShape = page.Drop(linkMaster, insertionPoint.X, insertionPoint.Y);
                WorkItemShapeInfo shapeInfo = new WorkItemShapeInfo(newWorkItemShape);
                shapeInfo.InitializeFromWorkItemRevision(workItem, int.Parse(relatedShape.ID));

                //int bugs = await VNC.AZDO.Helper.QueryRelatedBugsById(shapeInfo.Organization, int.Parse(shapeInfo.ID));

                //shapeInfo.RelatedBugs = bugs.ToString();

                shapeInfo.PopulateShapeDataFromInfo(newWorkItemShape, version);
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToDebugWindow($"{workItem.Id} - {ex}");
            }
        }

        private static Point GetPosition(MSVisio.Shape shape)
        {
            double x = 5.5;
            double y = 2.0;

            x = shape.CellsU["PinX"].ResultIU;
            y = shape.CellsU["PinY"].ResultIU;

            Point currentPosition = new Point(x, y);

            return currentPosition;
        }

        private static async Task<IList<WorkItem>> GetInfoById(WorkItemShapeInfo shapeInfo)
        {
            IList<WorkItem> result = null;
            int id = 0;

            if (!int.TryParse(shapeInfo.ID, out id))
            {
                MessageBox.Show($"Invalid WorkItem ID: ({shapeInfo.ID})");
            }
            else
            {
                result = await VNC.AZDO.Helper.QueryWorkItemInfoById(shapeInfo.Organization, int.Parse(shapeInfo.ID));

                int bugs = await VNC.AZDO.Helper.QueryRelatedBugsById(shapeInfo.Organization, int.Parse(shapeInfo.ID));
            }

            return result;
        }

        private static bool IsValidTeamProject(string organization, string teamProject)
        {
            // TODO(crhodes)
            // Go see if this is a valid Team Project
            return true;
        }

        private static async Task<IList<WorkItem>> GetInfoByTeamProject(WorkItemShapeInfo shapeInfo)
        {
            IList<WorkItem> result = null;

            string teamProject = shapeInfo.TeamProject;
            string workItemType = shapeInfo.WorkItemType;
            string state = shapeInfo.State;
            string areaPath = shapeInfo.AreaPath;
            string iterationPath = shapeInfo.IterationPath;

            if (!IsValidTeamProject(shapeInfo.Organization, teamProject))
            {
                MessageBox.Show($"Invalid TeamProject: ({teamProject})");
            }
            else
            {
                try
                {
                     result = await VNC.AZDO.Helper.QueryWorkItemInfoByTeam(
                         shapeInfo.Organization, teamProject, workItemType, state, areaPath, iterationPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error TeamProject: ({teamProject} ex:{ex})");
                }
            }

            return result;
        }

        public static async void QueryWorkItems(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] array)
        {
            Int64 startTicks = Log.APPLICATION("Enter", Common.LOG_CATEGORY);

            if (!VerifyRequiredStencils(app))
            {
                MessageBox.Show($"Cannot locate or open required Stencils, aborting.  Review Log for details");
                return;
            }

            MSVisio.Page activePage = app.ActivePage;
            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];

            WorkItemShapeInfo shapeInfo = new WorkItemShapeInfo(activeShape);

            // TODO(crhodes)
            // Logic here to decide what query to perform.
            // For now we support
            // TeamProject
            // TeamProject + WorkItemType
            // WorkItemType
            // ID

            IList<WorkItem> result = null;

            if (! string.IsNullOrEmpty(shapeInfo.TeamProject))
            {
                result = await GetInfoByTeamProject(shapeInfo);
            }
            else if (!string.IsNullOrEmpty(shapeInfo.ID))
            {
                result = await GetInfoById(shapeInfo);
            }

            if (result is null) return;

            if (result.Count > 0)
            {
                Point initialPosition = GetPosition(activeShape);
                Point insertionPoint = initialPosition;

                string stencilName = "Azure DevOps.vssx";

                MSVisio.Document linkStencil;
                MSVisio.Master linkMaster = null;
                string targetShapeName = activeShape.CellsU["Prop.WIShapeName"].ResultStrU[MSVisio.VisUnitCodes.visUnitsString];
                var version = WorkItemShapeInfo.WorkItemShapeVersion.V2;


                try
                {
                    linkStencil = app.Documents[stencilName];

                    try
                    {
                        linkMaster = linkStencil.Masters[targetShapeName];
                    }
                    catch (Exception ex)
                    {
                        VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("  Cannot find Master named:>{0}<", targetShapeName));
                    }
                }
                catch (Exception ex)
                {
                    VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("  Cannot find open Stencil named:>{0}<", stencilName));
                }

                // TODO(crhodes)
                // Figure out how to get size of shape from master.
                // HACK(crhodes)
                // .25 is for Link counts

                double height = version == WorkItemShapeInfo.WorkItemShapeVersion.V1 ? 0.375 : 0.475;

                WorkItemOffsets workItemOffsets = new WorkItemOffsets(initialPosition, height: height, padX: 0.25, padY: 0.05);

                foreach (var linkedWorkItem in result)
                {
                    //// NOTE(crhodes)
                    //// This includes the current shape.  Do not add it.
                    //// May always be first one.  Maybe loop counter
                    //if (linkedWorkItem.Id == id)
                    //{
                    //    continue;
                    //}

                    VNCVisioAddIn.Common.WriteToDebugWindow($"{linkedWorkItem.Id} {linkedWorkItem.Fields["System.Title"]}");

                    insertionPoint = AZDOPageLayout.CalculateInsertionPointQueriedWorkItems(initialPosition, linkedWorkItem, shapeInfo, workItemOffsets);

                    AddNewWorkItemShapeToPage(activePage, linkMaster, linkedWorkItem, insertionPoint, shapeInfo, version);
                }
            }

            Log.APPLICATION("Exit", Common.LOG_CATEGORY, startTicks);
        }
        
        public static async void AddLinkedWorkItemsExternal(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] array)
        {

            MSVisio.Page activePage = app.ActivePage;
            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
            var version = WorkItemShapeInfo.WorkItemShapeVersion.V2;

            WorkItemShapeInfo activeShapeWorkItemInfo = new WorkItemShapeInfo(activeShape);

            int id;

            if (int.TryParse(activeShapeWorkItemInfo.ID, out id))
            {
            }
            else
            {
                MessageBox.Show($"Cannot parse ({activeShapeWorkItemInfo.ID}) as WorkItemID");
                return;
            }

            int relatedLinkCount;

            if (int.TryParse(activeShapeWorkItemInfo.RelatedLinkCount, out relatedLinkCount))
            {
            }
            else
            {
                MessageBox.Show($"Cannot parse ({activeShapeWorkItemInfo.RelatedLinkCount}) as RelatedLinkCount");
                return;
            }

            var result = await VNC.AZDO.Helper.QueryWorkItemLinks(activeShapeWorkItemInfo.Organization, id, relatedLinkCount);

            if (result.Count > 0)
            {
                Point initialPosition = GetPosition(activeShape);
                Point insertionPoint = initialPosition;

                string stencilName = "Azure DevOps.vssx";

                string targetShapeName = activeShape.CellsU["Prop.WIShapeName"].ResultStrU[MSVisio.VisUnitCodes.visUnitsString];
                //string shapeName = "WI 2";

                MSVisio.Document linkStencil;
                MSVisio.Master linkMaster = null;

                try
                {
                    linkStencil = app.Documents[stencilName];

                    try
                    {
                        linkMaster = linkStencil.Masters[targetShapeName];
                    }
                    catch (Exception ex)
                    {
                        VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("  Cannot find Master named:>{0}<", targetShapeName));
                    }
                }
                catch (Exception ex)
                {
                    VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("  Cannot find open Stencil named:>{0}<", stencilName));
                }

                // TODO(crhodes)
                // Figure out how to get size of shape from master.
                // HACK(crhodes)
                // .25 is for Link counts

                double height = version == WorkItemShapeInfo.WorkItemShapeVersion.V1 ? 0.375 : 0.475;

                WorkItemOffsets workItemOffsets = new WorkItemOffsets(initialPosition, height: height, padX: 0.25, padY: 0.05);

                foreach (var linkedWorkItem in result)
                {
                    // NOTE(crhodes)
                    // This includes the current shape.  Do not add it.
                    // May always be first one.  Maybe loop counter
                    if (linkedWorkItem.Id == id)
                    {
                        continue;
                    }

                    VNCVisioAddIn.Common.WriteToDebugWindow($"{linkedWorkItem.Id} {linkedWorkItem.Fields["System.Title"]}");

                    insertionPoint = AZDOPageLayout.CalculateInsertionPointLinkedWorkItems(initialPosition, linkedWorkItem, activeShapeWorkItemInfo, workItemOffsets);

                    AddNewWorkItemShapeToPage(activePage, linkMaster, linkedWorkItem, insertionPoint, activeShapeWorkItemInfo, version);
                }
            }

            VNCVisioAddIn.Common.WriteToDebugWindow($"{activeShapeWorkItemInfo}");
        }

        public static async void GetWorkItemRevisions(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] array)
        {
            VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}()",
                MethodBase.GetCurrentMethod().Name));

            // NOTE(crhodes)
            // Can launch a UI here.  Or earlier.

            //DxThemedWindowHost.DisplayUserControlInHost(ref addLinkedWorkItemsHost,
            //    "Edit Shape Control Points Text",
            //    Common.DEFAULT_WINDOW_WIDTH, Common.DEFAULT_WINDOW_HEIGHT,
            //    DxThemedWindowHost.ShowWindowMode.Modeless,
            //    new Presentation.Views.EditControlPoints());

            MSVisio.Page activePage = app.ActivePage;
            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
            string targetShapeName = activeShape.CellsU["Prop.WIShapeName"].ResultStrU[MSVisio.VisUnitCodes.visUnitsString];
            var version = WorkItemShapeInfo.WorkItemShapeVersion.V2;

            //AddLinkedWorkItems(app, activePage, activeShape, targetShapeName, version);

            if (!VerifyRequiredStencils(app))
            {
                return;
            }

            WorkItemShapeInfo activeShapeWorkItemInfo = new WorkItemShapeInfo(activeShape);

            int id;

            if (int.TryParse(activeShapeWorkItemInfo.ID, out id))
            {
            }
            else
            {
                MessageBox.Show($"Cannot parse ({activeShapeWorkItemInfo.ID}) as WorkItemID");
                return;
            }

            //int relatedLinkCount;

            //if (int.TryParse(activeShapeWorkItemInfo.RelatedLinkCount, out relatedLinkCount))
            //{
            //}
            //else
            //{
            //    MessageBox.Show($"Cannot parse ({activeShapeWorkItemInfo.RelatedLinkCount}) as RelatedLinkCount");
            //    return;
            //}

            //var result = await VNC.AZDO.Helper.QueryWorkItemLinks(activeShapeWorkItemInfo.Organization, id, relatedLinkCount);

            var result = await VNC.AZDO.Helper.QueryWorkItemRevisionsById(activeShapeWorkItemInfo.Organization, id);

            if (result.Count > 0)
            {
                Point initialPosition = GetPosition(activeShape);
                Point insertionPoint = initialPosition;

                string stencilName = "Azure DevOps.vssx";

                MSVisio.Document linkStencil;
                MSVisio.Master linkMaster = null;

                try
                {
                    linkStencil = app.Documents[stencilName];

                    try
                    {
                        linkMaster = linkStencil.Masters[targetShapeName];
                    }
                    catch (Exception ex)
                    {
                        VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("  Cannot find Master named:>{0}<", targetShapeName));
                    }
                }
                catch (Exception ex)
                {
                    VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("  Cannot find open Stencil named:>{0}<", stencilName));
                }

                // TODO(crhodes)
                // Figure out how to get size of shape from master.
                // HACK(crhodes)
                // .25 is for Link counts

                double height = version == WorkItemShapeInfo.WorkItemShapeVersion.V1 ? 0.375 : 0.475;

                WorkItemOffsets workItemOffsets = new WorkItemOffsets(initialPosition, height: height, padX: 0.25, padY: 0.05);

                foreach (var linkedWorkItem in result)
                {
                    //// NOTE(crhodes)
                    //// This includes the current shape.  Do not add it.
                    //// May always be first one.  Maybe loop counter
                    //if (linkedWorkItem.Id == id)
                    //{
                    //    continue;
                    //}

                    VNCVisioAddIn.Common.WriteToDebugWindow($"{linkedWorkItem.Id} {linkedWorkItem.Fields["System.Title"]}");

                    insertionPoint = AZDOPageLayout.CalculateInsertionPointLinkedWorkItems(initialPosition, linkedWorkItem, activeShapeWorkItemInfo, workItemOffsets);

                    AddNewWorkItemRevisionShapeToPage(activePage, linkMaster, linkedWorkItem, insertionPoint, activeShapeWorkItemInfo, version);
                }
            }

            //var result = await VNC.AZDO.Helper.QueryWorkItemRevisionsById(shapeInfo.Organization, id);
        }
    }
}