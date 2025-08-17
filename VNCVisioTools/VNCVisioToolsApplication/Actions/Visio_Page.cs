using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;

//using EnvDTE;

using VNC.Core;


using VNC;

using MSVisio = Microsoft.Office.Interop.Visio;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Actions
{
    public class Visio_Page
    {
        #region Enums, Fields, Properties, Structures

        private enum LayerNameType
        {
            AllNames = 0,
            AddName = 1,
            RemovalName = 2
        }

        // Replace all occurrences of:
        // Common.WriteToDebugWindow(string.Format("{0}()", MethodBase.GetCurrentMethod().Name));
        // with:
        // Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

        public static void AddDefaultLayers()
        {
            VNC.Log.TRACE("", Common.LOG_CATEGORY, 0);

            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;
            MSVisio.Document doc = app.ActiveDocument;
            MSVisio.Page page = app.ActivePage;

            AddDefaultLayers(page);
        }

        public static void AddDefaultLayers(MSVisio.Page page)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}({page.NameU})");

            if (page == null)
            {
                System.Windows.Forms.MessageBox.Show("No ActivePage");
                return;
            }

            MSVisio.Layers layers = page.Layers;

            try
            {
                MSVisio.Page layersPage = Common.VisioApplication.ActiveDocument.Pages["Default Layers"];
                //Common.WriteToDebugWindow(string.Format("  Copying {0} links", linkPage.Shapes.Count));

                List<MSVisio.Shape> layerNames = GetLayerNameShapes(layersPage, LayerNameType.AddName);

                string layerName = null;

                // These are the defaults if the shape does not have values.

                string layerVisible = "1";
                string layerPrint = "1";
                string layerActive = "0";
                string layerLock = "0";
                string layerSnap = "1";
                string layerGlue = "1";

                foreach (MSVisio.Shape shape in layerNames)
                {
                    AddLayer(page, shape.Text, layerVisible, layerPrint, layerActive, layerLock, layerSnap, layerGlue);
                }
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force:true);
                // No navigation Links Page perhaps
            }
        }

        public static MSVisio.Layer AddLayer(MSVisio.Page page,
                    string layerName,
            string layerVisible = "1",
            string layerPrint = "1",
            string layerActive = "0",
            string layerLock = "0",
            string layerSnap = "1",
            string layerGlue = "1")
        {
            MSVisio.Layer layer = null;

            try
            {
                if (page.Layers.Count > 0)
                {
                    // See if layer already exists

                    try
                    {
                        layer = page.Layers[layerName];
                    }
                    catch (Exception ex)
                    {

                    }
                }

                if (layer == null)
                {
                    layer = page.Layers.Add(layerName);
                }

                layer.CellsC[(short)MSVisio.VisCellIndices.visLayerVisible].FormulaU = layerVisible;

                layer.CellsC[(short)MSVisio.VisCellIndices.visLayerPrint].FormulaU = layerPrint;

                layer.CellsC[(short)MSVisio.VisCellIndices.visLayerActive].FormulaU = layerActive;

                layer.CellsC[(short)MSVisio.VisCellIndices.visLayerLock].FormulaU = layerLock;

                layer.CellsC[(short)MSVisio.VisCellIndices.visLayerSnap].FormulaU = layerSnap;

                layer.CellsC[(short)MSVisio.VisCellIndices.visLayerGlue].FormulaU = layerGlue;
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force:true);
            }

            return layer;
        }

        public static void AddNavigationLinks(MSVisio.Page page)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}({page.NameU})");

            if ((page.Background != 0) || (page.NameU == "Navigation Links"))
            {
                //Common.WriteToDebugWindow("   Skipping");
                return;
            }

            RemoveNavigationLinks(page);

            //List<MSVisio.Shape> links = Actions.Visio_Document.GetNavigationLinks();

            try
            {
                MSVisio.Window activeWindow = Common.VisioApplication.ActiveWindow;
                activeWindow.Page = Common.VisioApplication.ActiveDocument.Pages["Navigation Links"];
                activeWindow.SelectAll();
                activeWindow.Selection.Copy(MSVisio.VisCutCopyPasteCodes.visCopyPasteNoTranslate);
                activeWindow.Page = page;
                page.Paste(MSVisio.VisCutCopyPasteCodes.visCopyPasteNoTranslate);

                //Common.VisioApplication.Windows.ItemEx["Navigation Links"].Activate();
                //Common.VisioApplication.ActiveWindow.SelectAll();
                //Common.VisioApplication.ActiveWindow.Selection.Copy();
                //Common.VisioApplication.Windows.ItemEx["Navigation Links"].Activate();


                //MSVisio.Page linkPage = Common.VisioApplication.ActiveDocument.Pages["Navigation Links"];
                //linkPage.Application.
                //Common.VisioApplication.
                //Common.WriteToDebugWindow(string.Format("  Copying {0} links", linkPage.Shapes.Count));

                //foreach (MSVisio.Shape shape in linkPage.Shapes)
                //{
                //    // TODO: Make this smarter about only using IsNavigationLink shapes
                //    shape.Copy(MSVisio.VisCutCopyPasteCodes.visCopyPasteNoTranslate);
                //    page.Paste(MSVisio.VisCutCopyPasteCodes.visCopyPasteNoTranslate);
                //}

                // Typically we don't print the stuff on the navigation layer.

                page.Layers["Navigation"].CellsC[(short)MSVisio.VisCellIndices.visLayerPrint].FormulaU = "0";
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force:true);
                // No navigation Links Page perhaps
            }
        }

        public static void AutoSizePageOff()
        {
            VNC.Log.TRACE("", Common.LOG_CATEGORY, 0);

            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;
            MSVisio.Document doc = app.ActiveDocument;
            MSVisio.Page page = app.ActivePage;

            AutoSizePageOff(page);
        }

        public static void AutoSizePageOff(MSVisio.Page page)
        {
            page.AutoSize = false;
        }

        public static void AutoSizePageOn()
        {
            VNC.Log.TRACE("", Common.LOG_CATEGORY, 0);

            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;
            MSVisio.Document doc = app.ActiveDocument;
            MSVisio.Page page = app.ActivePage;

            AutoSizePageOn(page);
        }
        public static void AutoSizePageOn(MSVisio.Page page)
        {
            page.AutoSize = true;
        }
        public static void CreateActivityPage(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] args)
        {
            string pageLevel = null;
            string backgroundPageName = null;

            if (args.Count() != 2)
            {
                Common.WriteToDebugWindow(string.Format("Incorrect Argument Count, expected 2.  Check ShapeSheet"));
            }
            else
            {
                pageLevel = args[0];
                backgroundPageName = args[1];
            }

            Common.WriteToDebugWindow(string.Format("{0}() PageLevel:{1}  Background:{2}",
                MethodBase.GetCurrentMethod().Name,
                pageLevel, backgroundPageName));

            // Current shape contains text for new page name.

            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
            Common.WriteToDebugWindow(string.Format("  Shape(Name:{0}  Text:{1}", activeShape.Name, activeShape.Text));

            //string newPageName = string.Format("{0}-{1}", pageLevel, activeShape.Text);
            string newPageName = string.Format("{0}{1}{2}", pageLevel, "-", activeShape.Text);

            MSVisio.Page newPage = CreatePage(newPageName, backgroundPageName);

            // Update the current shape's hyperlink to point to the new page

            // TODO(crhodes):
            //	Not sure which of these two approaches is doing the magic.

            MSVisio.Hyperlink currentHyperLink = activeShape.AddHyperlink();
            currentHyperLink.SubAddress = newPageName;

            //activeShape.CellsSRC[(short)MSVisio.VisSectionIndices.visSectionHyperlink, 0, 2].Formula = newPageName;


            // TODO(crhodes): 
            // Add User Section data depending on pageLevel argument, e.g. L0, L1, L2, ...

            switch (pageLevel)
            {
                case "L0":

                    break;
                case "L1":

                    break;

                case "L2":

                    break;

                default:

                    break;
            }

            //  Figure out how to get PageName shape added.
        }

        public static void CreateArtifactPage(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] args)
        {
            string pageLevel = null;
            string backgroundPageName = null;

            if (args.Count() != 2)
            {
                Common.WriteToDebugWindow(string.Format("Incorrect Argument Count, expected 2.  Check ShapeSheet"));
            }
            else
            {
                pageLevel = args[0];
                backgroundPageName = args[1];
            }

            Common.WriteToDebugWindow(string.Format("{0}() PageLevel:{1}  Background:{2}",
                MethodBase.GetCurrentMethod().Name,
                pageLevel, backgroundPageName));

            // Current shape contains text for new page name.

            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
            Common.WriteToDebugWindow(string.Format("  Shape(Name:{0}  Text:{1}", activeShape.Name, activeShape.Text));

            //string newPageName = string.Format("{0}-{1}", pageLevel, activeShape.Text);
            string newPageName = string.Format("{0}{1}{2}", "", "", activeShape.Text);

            MSVisio.Page newPage = CreatePage(newPageName, backgroundPageName);

            // Update the current shape's hyperlink to point to the new page

            // TODO(crhodes):
            //	Not sure which of these two approaches is doing the magic.

            MSVisio.Hyperlink currentHyperLink = activeShape.AddHyperlink();
            currentHyperLink.SubAddress = newPageName;

            activeShape.CellsSRC[(short)MSVisio.VisSectionIndices.visSectionHyperlink, 0, 2].Formula = newPageName;

            //  Figure out how to get PageName shape added.
        }

        public static void CreateDefaultLayersPage(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] array)
        {
            MSVisio.Page newPage = CreatePage(pageName: "Default Layers", backgroundPageName: "", isBackground: 1);
        }

        public static void CreateMetricPage(MSVisio.Application app, string doc, string page, string shape, string shapeu, String[] args)
        {
            string arg0 = null;
            string backgroundPageName = null;

            if (args.Count() != 2)
            {
                Common.WriteToDebugWindow(string.Format("Incorrect Argument Count, expected 2.  Check ShapeSheet"));
            }
            else
            {
                arg0 = args[0];
                backgroundPageName = args[1];
            }

            Common.WriteToDebugWindow(string.Format("{0}() arg0:{1}  Background:{2}",
                MethodBase.GetCurrentMethod().Name,
                arg0, backgroundPageName));

            // Current shape contains text for new page name.

            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
            Common.WriteToDebugWindow(string.Format("  Shape(Name:{0}  Text:{1}", activeShape.Name, activeShape.Text));

            string newPageName = string.Format("{0}{1}{2}", "", "", activeShape.Text);

            MSVisio.Page newPage = CreatePage(newPageName, backgroundPageName);

            MSVisio.Hyperlink currentHyperLink = activeShape.AddHyperlink();
            currentHyperLink.SubAddress = newPageName;

            activeShape.CellsSRC[(short)MSVisio.VisSectionIndices.visSectionHyperlink, 0, 2].Formula = newPageName;

            // TODO(crhodes):
            //	Figure out what to do with roleSource
            //  Figure out how to get PageName shape added.
        }

        public static void CreateNavigationLinksPage(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] array)
        {
            MSVisio.Page newPage = CreatePage(pageName: "Navigation Links", backgroundPageName: "", isBackground: 1);
        }

        public static MSVisio.Page CreatePage(string pageName, string backgroundPageName, short isBackground = 0)
        {
            Common.WriteToDebugWindow(string.Format("{0}() Page:{1}  Background:{2}",
                MethodBase.GetCurrentMethod().Name,
                pageName, backgroundPageName));

            // TODO(crhodes):
            //	Error handling. Page already exists, background page doesn't exist, etc.
            MSVisio.Application app = Common.VisioApplication;
            int currentPageIndex = app.ActivePage.Index;

            Common.WriteToDebugWindow(string.Format("   currentPageIndex:{0}", currentPageIndex));
            MSVisio.Page newPage = app.ActiveDocument.Pages.Add();

            // Cleanup page names
            pageName = pageName.Replace("\n", " ");

            newPage.Name = pageName;

            try
            {
                newPage.BackPage = backgroundPageName;
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(string.Format("Cannot Find Background Page ({0})", backgroundPageName));
            }

            newPage.Index = (short)(currentPageIndex + 1);

            newPage.Background = isBackground;

            AddNavigationLinks(newPage);

            return newPage;
        }

        public static void CreatePageBasePage(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] array)
        {
            MSVisio.Page newPage = CreatePage(pageName: "Page Base", backgroundPageName: "", isBackground: 1);
        }

        public static void CreatePageForShape(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] args)
        {
            string prefix = null;
            string delimiter = null;
            string backgroundPageName = null;

            if (args.Count() != 3)
            {
                Common.WriteToDebugWindow(string.Format("Incorrect Argument Count, expected 3.  Check ShapeSheet"));
            }
            else
            {
                prefix = args[0];
                delimiter = args[1];
                backgroundPageName = args[2];
            }

            Common.WriteToDebugWindow(string.Format("{0}() prefix:>{1}< delimiter:>{2}< backgroundPageName:>{3}<",
                MethodBase.GetCurrentMethod().Name,
                prefix, delimiter, backgroundPageName));

            try
            {
                // Current shape contains text for new page name.
                MSVisio.Page activePage = app.ActivePage;
                MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
                Common.WriteToDebugWindow(string.Format("  Shape(Name:>{0}< Text:>{1}< Characters:>{2}<", activeShape.Name, activeShape.Text, activeShape.Characters.TextAsString));

                string shapePageName = "Error-PageNameNotProvided";

                if (activeShape.CellExistsU["Prop.PageName", 0] != 0)
                {
                    shapePageName = activeShape.CellsU["Prop.PageName"].ResultStrU[MSVisio.VisUnitCodes.visUnitsString];
                }
                else if (activeShape.Characters.TextAsString.Length > 0)
                {
                    //string newPageName = string.Format("{0}{1}{2}", prefix, delimiter, activeShape.Text);
                    // shape.Text comes in as OBJ if use fields and Shape Data.   Use shape.Characters instead.
                    shapePageName = activeShape.Characters.TextAsString;
                }

                string newPageName = string.Format("{0}{1}{2}", prefix, delimiter, shapePageName);

                MSVisio.Page newPage = CreatePage(newPageName, backgroundPageName);

                // The old style linkable masters did not have Prop.Data for the HyperLink.  Check first before updating.
                // Should really retire all the old shapes and remove this code.

                if (activeShape.CellExistsU["Prop.HyperLink", 0] != 0)
                {
                    activeShape.CellsU["Prop.HyperLink"].FormulaU = newPageName.WrapInDblQuotes();
                }
                else
                {
                    MSVisio.Hyperlink currentHyperLink = activeShape.AddHyperlink();
                    currentHyperLink.SubAddress = newPageName;
                }

                // Check to see if there is a ReturnLink Property with values that can be used to create a return link
                // to the page that linked to us.

                if (activeShape.CellExistsU["Prop.ReturnLink", 0] != 0)
                {
                    //string returnLinkProp = activeShape.CellsU["Prop.ReturnLink"].FormulaU;   // This returns "<string>"  we want just <string>
                    string returnLinkProp = activeShape.CellsU["Prop.ReturnLink"].ResultStrU[MSVisio.VisUnitCodes.visUnitsString];
                    string[] linkInfo = returnLinkProp.Split(',');
                    string stencilName = linkInfo[0];
                    string shapeName = linkInfo[1];

                    Common.WriteToDebugWindow(string.Format("  returnLinkProp:>{0}< stencilName:>{1}< shapeName:>{2}< ", returnLinkProp, stencilName, shapeName));

                    try
                    {
                        MSVisio.Document linkStencil = app.Documents[stencilName];

                        try
                        {
                            MSVisio.Master linkMaster = linkStencil.Masters[shapeName];

                            // Add return link in upper left corner.  Assume 11x8.5

                            // TODO(crhodes)
                            // Get Page Size and drop in upper left
                            MSVisio.Shape returnLinkShape = newPage.Drop(linkMaster, 1.0, 8.0);

                            returnLinkShape.CellsU["Prop.PageName"].FormulaU = activePage.Name.WrapInDblQuotes();
                            returnLinkShape.CellsU["Prop.HyperLink"].FormulaU = activePage.Name.WrapInDblQuotes();
                        }
                        catch (Exception ex)
                        {
                            Common.WriteToDebugWindow(string.Format("  Cannot find Master named:>{0}<", shapeName));
                        }
                    }
                    catch (Exception ex)
                    {
                        Common.WriteToDebugWindow(string.Format("  Cannot find open Stencil named:>{0}<", stencilName));
                    }
                }

                // Add a header.  May want to pick the stencil and shape for config file.
                // Or add a property to Shape.

                VNCVisioAddIn.Helpers.LoadStencil(app, "Page Shapes.vssx");
                MSVisio.Master headerMaster = app.Documents[@"Page Shapes.vssx"].Masters[@"18pt Header"];

                newPage.Drop(headerMaster, 5.5, 8.0625);

                // NOTE(crhodes)
                // Add the shape that triggered the event.  User can delete if doesn't want.
                // More and more often I go back and copy it, traverse the link, and drop it.
                // Drop in middle of page for now assuming 11x8.5

                // TODO(crhodes)
                // Get Page Size and drop in center
                newPage.Drop(activeShape, 5.5, 4.25);
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force:true);
            }
        }

        public static void CreateRolePage(MSVisio.Application app, string doc, string page, string shape, string shapeu, String[] args)
        {
            string roleSource = null;
            string backgroundPageName = null;

            if (args.Count() != 2)
            {
                Common.WriteToDebugWindow(string.Format("Incorrect Argument Count, expected 2.  Check ShapeSheet"));
            }
            else
            {
                roleSource = args[0];
                backgroundPageName = args[1];
            }

            Common.WriteToDebugWindow(string.Format("{0}() PageLevel:{1}  Background:{2}",
                MethodBase.GetCurrentMethod().Name,
                roleSource, backgroundPageName));

            // Current shape contains text for new page name.

            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
            Common.WriteToDebugWindow(string.Format("  Shape(Name:{0}  Text:{1}", activeShape.Name, activeShape.Text));

            string newPageName = string.Format("{0}{1}{2}", "", "", activeShape.Text);

            MSVisio.Page newPage = CreatePage(newPageName, backgroundPageName);

            MSVisio.Hyperlink currentHyperLink = activeShape.AddHyperlink();
            currentHyperLink.SubAddress = newPageName;

            activeShape.CellsSRC[(short)MSVisio.VisSectionIndices.visSectionHyperlink, 0, 2].Formula = newPageName;

            // TODO(crhodes):
            //	Figure out what to do with roleSource
            //  Figure out how to get PageName shape added.
        }

        public static void CreateToolPage(MSVisio.Application app, string doc, string page, string shape, string shapeu, String[] args)
        {
            string arg0 = null;
            string backgroundPageName = null;

            if (args.Count() != 2)
            {
                Common.WriteToDebugWindow(string.Format("Incorrect Argument Count, expected 2.  Check ShapeSheet"));
            }
            else
            {
                arg0 = args[0];
                backgroundPageName = args[1];
            }

            Common.WriteToDebugWindow(string.Format("{0}() arg0:{1}  Background:{2}",
                MethodBase.GetCurrentMethod().Name,
                arg0, backgroundPageName));

            // Current shape contains text for new page name.

            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
            Common.WriteToDebugWindow(string.Format("  Shape(Name:{0}  Text:{1}", activeShape.Name, activeShape.Text));

            string newPageName = string.Format("{0}{1}{2}", "", "", activeShape.Text);

            MSVisio.Page newPage = CreatePage(newPageName, backgroundPageName);

            MSVisio.Hyperlink currentHyperLink = activeShape.AddHyperlink();
            currentHyperLink.SubAddress = newPageName;

            activeShape.CellsSRC[(short)MSVisio.VisSectionIndices.visSectionHyperlink, 0, 2].Formula = newPageName;

            // TODO(crhodes):
            //  Figure out how to get PageName shape added.
        }

        public static void DeleteLayer(MSVisio.Page page, string layerName, short deleteShapes)
        {
            // TODO(crhodes)
            // may want to pass in a forceUnlock flag defaulted to 0

            try
            {
                MSVisio.Layer layer = null;

                if (page.Layers.Count > 0)
                {
                    // See if layer already exists

                    try
                    {
                        layer = page.Layers[layerName];
                    }
                    catch (Exception ex)
                    {

                    }
                }

                if (layer != null)
                {
                    layer.CellsC[(short)MSVisio.VisCellIndices.visLayerLock].FormulaU = "0";
                    layer.Delete(deleteShapes);
                }
                else
                {
                    Common.WriteToDebugWindow(string.Format("Layer >{0}< does not exist", layerName));
                }
            }
            catch (Exception ex)
            {
                // TODO(crhodes):
                // Decide if what to show this to user.  Layer maybe locked.
                Common.WriteToDebugWindow(ex.ToString(), force:true);
            }
        }

        public static void DeleteLayers(MSVisio.Page page)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name} Name:>{page.Name}< NameU:>{page.NameU}<");

            try
            {
                // TODO(crhodes):
                // Handle if "Default Layers" page doesn't exist

                MSVisio.Page layersPage = Common.VisioApplication.ActiveDocument.Pages["Default Layers"];
                List<MSVisio.Shape> layerNames = GetLayerNameShapes(layersPage, LayerNameType.RemovalName);

                foreach (MSVisio.Shape shape in layerNames)
                {
                    DeleteLayer(page, shape.Text, 0);
                    //foreach (MSVisio.Layer layer in page.Layers)
                    //{
                    //    if (layer.NameU.Equals(shape.Text))
                    //    {
                    //        layer.Delete(0);
                    //    }
                    //}
                }
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force:true);
            }
        }

        public static void DisplayInfo(MSVisio.Page page)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            if (page == null)
            {
                MessageBox.Show("No ActivePage");
                return;
            }
            Common.WriteToDebugWindow($"ActivePage.Name - {page.Name}\n");
            Common.WriteToDebugWindow($"ActivePage.NameU - {page.NameU}\n");

            try
            {
                Common.WriteToDebugWindow($"ActivePage.OriginalPage.Name - {page.OriginalPage.Name}\n");
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow($"ActivePage.OriginalPage.Name - <none>\n");
            }

            Common.WriteToDebugWindow($"ActivePage.AutoSize - {page.AutoSize}\n");
            Common.WriteToDebugWindow($"ActivePage.Background - {page.Background}\n");
            Common.WriteToDebugWindow($"ActivePage.Comments - {page.Comments.ToString()}\n");
            Common.WriteToDebugWindow($"ActivePage.Connects - {page.Connects.Count}\n");
            Common.WriteToDebugWindow($"ActivePage.ID - {page.ID}\n");
            Common.WriteToDebugWindow($"ActivePage.Index - {page.Index}\n");
            Common.WriteToDebugWindow($"ActivePage.Layers - {page.Layers.Count}\n");
            Common.WriteToDebugWindow($"ActivePage.LayoutRoutePassive - {page.LayoutRoutePassive}\n");

            Common.WriteToDebugWindow($"ActivePage.PageSheet.Name - {page.PageSheet.Name}\n");
            Common.WriteToDebugWindow($"ActivePage.PrintTileCount - {page.PrintTileCount}\n");

            try
            {
                Common.WriteToDebugWindow($"ActivePage.ReviewerID - {page.ReviewerID}\n");
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow($"ActivePage.ReviewerID - <none>\n");
            }

            Common.WriteToDebugWindow($"ActivePage.ShapeComments - {page.ShapeComments.ToString()}\n");
            Common.WriteToDebugWindow($"ActivePage.Shapes - {page.Shapes.Count}\n");
            Common.WriteToDebugWindow($"ActivePage.Stat - {page.Stat}\n");
            Common.WriteToDebugWindow($"ActivePage.Type - {page.Type.ToString()}\n");

            foreach (MSVisio.Shape shape in page.Shapes)
            {
                Visio_Shape.DisplayInfo(shape);
            }
        }

        public static void DisplayLayer(MSVisio.Page page, string layerName, bool show)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}(layer:{layerName} show:{show})");

            Common.WriteToDebugWindow(page.NameU);

            foreach (MSVisio.Layer layer in page.Layers)
            {
                Common.WriteToDebugWindow($"  {layer.Name} - Visible:{layer.CellsC[(short)MSVisio.VisCellIndices.visLayerVisible].FormulaU.ToString()} Print:{layer.CellsC[(short)MSVisio.VisCellIndices.visLayerPrint].FormulaU.ToString()}");

                if (layer.Name == layerName)
                {
                    layer.CellsC[(short)MSVisio.VisCellIndices.visLayerVisible].FormulaU = (show == true ? "1" : "0");
                }
            }
        }
        public static void LockLayer(string layerName)
        {
            Common.WriteToDebugWindow(string.Format("{0}()",
                MethodBase.GetCurrentMethod().Name));

            MSVisio.Layer layer = Common.VisioApplication.ActivePage.Layers[layerName];

            layer.CellsC[(short)MSVisio.VisCellIndices.visLayerLock].Formula = "1";
        }

        public static void MovePage(MSVisio.Page page, string targetDocument)
        {
            Common.WriteToDebugWindow(string.Format("{0}()",
                MethodBase.GetCurrentMethod().Name));
            MSVisio.Application app = page.Application;
            MSVisio.Document doc = page.Document;
            string currentDocument = doc.Name;

            try
            {
                Int32 diagramServices = doc.DiagramServicesEnabled;
                doc.DiagramServicesEnabled = (Int32)MSVisio.VisDiagramServices.visServiceAll;

                app.ActiveWindow.Page = page;

                // TODO(crhodes)
                // Need to get all the shapes on the page and ignore the Navigation Shapes

                app.ActiveWindow.SelectAll();
                app.ActiveWindow.Selection.Cut();
                app.Windows.ItemEx[targetDocument].Activate();

                MSVisio.Page newPage = app.ActiveDocument.Pages.Add();
                newPage.Name = page.Name;
                newPage.Paste();

                // Add navigation links from target document - 
                //  will automatically remove current ones from source document

                AddNavigationLinks(newPage);

                // Return to original document and delete the page
                app.Windows.ItemEx[currentDocument].Activate();
                short renumberPages = 0;    // Do not renumber default named pages
                page.Delete(renumberPages);
            }
            catch (Exception ex)
            {
                // TODO(crhodes):
                // Decide if what to show this to user.  Layer maybe locked.
                Common.WriteToDebugWindow(ex.ToString(), force:true);
            }

            Common.WriteToDebugWindow("Exit");
        }

        public static void PageChanged(MSVisio.Page page)
        {
            Common.WriteToDebugWindow(string.Format("{0} Name:>{1}< NameU:>{2}<",
                MethodBase.GetCurrentMethod().Name, page.Name, page.NameU));

            SyncPageNames(page);
            UpdatePageNameShapes(page);
        }

        public static void PrintPage()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            try
            {
                MSVisio.Application app = Common.VisioApplication;
                MSVisio.Document doc = app.ActiveDocument;
                MSVisio.Page page = app.ActivePage;

                page.Print();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force:true);
            }
        }

        public static void RemoveLayers()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;
            MSVisio.Document doc = app.ActiveDocument;
            MSVisio.Page page = app.ActivePage;

            DeleteLayers(page);
        }

        public static void SavePage(MSVisio.Page page)
        {
            Common.WriteToDebugWindow(string.Format("{0} Name:>{1}< NameU:>{2}<",
                MethodBase.GetCurrentMethod().Name, page.Name, page.NameU));

            MSVisio.Application app = Common.VisioApplication;

            app.Settings.SetRasterExportResolution(MSVisio.VisRasterExportResolution.visRasterUseCustomResolution, 150, 150, MSVisio.VisRasterExportResolutionUnits.visRasterPixelsPerCm);
            //app.Settings.SetRasterExportResolution(MSVisio.VisRasterExportResolution.visRasterUseCustomResolution, 300, 300, MSVisio.VisRasterExportResolutionUnits.visRasterPixelsPerCm);
            app.Settings.SetRasterExportSize(MSVisio.VisRasterExportSize.visRasterFitToSourceSize);
            //app.Settings.SetRasterExportSize(MSVisio.VisRasterExportSize.visRasterFitToScreenSize);
            //app.Settings.SetRasterExportSize(MSVisio.VisRasterExportSize.visRasterFitToPrinterSize);
            //app.Settings.SetRasterExportSize(MSVisio.VisRasterExportSize.visRasterFitToCustomSize, 11.0, 8.5, MSVisio.VisRasterExportSizeUnits.visRasterInch);

            //app.Settings.SetRasterExportSize(MSVisio.VisRasterExportSize.visRasterFitToSourceSize, 9.56, 7.47, MSVisio.VisRasterExportSizeUnits.visRasterInch);
            app.Settings.RasterExportDataFormat = MSVisio.VisRasterExportDataFormat.visRasterInterlace;
            app.Settings.RasterExportColorFormat = MSVisio.VisRasterExportColorFormat.visRaster24Bit;
            app.Settings.RasterExportRotation = MSVisio.VisRasterExportRotation.visRasterNoRotation;
            app.Settings.RasterExportFlip = MSVisio.VisRasterExportFlip.visRasterNoFlip;
            app.Settings.RasterExportBackgroundColor = 16777215;
            app.Settings.RasterExportTransparencyColor = 16777215;
            app.Settings.RasterExportUseTransparencyColor = false;

            string pageName = GetPageSaveName(page);

            try
            {
                page.Export(pageName);
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force:true);
            }

            // From Macro Recorder
            //'Enable diagram services
            //Dim DiagramServices As Integer
            //DiagramServices = ActiveDocument.DiagramServicesEnabled
            //ActiveDocument.DiagramServicesEnabled = visServiceVersion140 + visServiceVersion150

            //Application.Settings.SetRasterExportResolution visRasterUseCustomResolution, 300#, 300#, visRasterPixelsPerInch
            //Application.Settings.SetRasterExportSize visRasterFitToSourceSize, 9.5625, 7.472222, visRasterInch
            //Application.Settings.RasterExportDataFormat = visRasterInterlace
            //Application.Settings.RasterExportColorFormat = visRaster24Bit
            //Application.Settings.RasterExportRotation = visRasterNoRotation
            //Application.Settings.RasterExportFlip = visRasterNoFlip
            //Application.Settings.RasterExportBackgroundColor = 16777215
            //Application.Settings.RasterExportTransparencyColor = 16777215
            //Application.Settings.RasterExportUseTransparencyColor = False


            //Application.ActiveWindow.Page.Export "C:\temp\TestDrawing2.png"

            //'Restore diagram services
            //ActiveDocument.DiagramServicesEnabled = DiagramServices



        }

        //End Sub
        public static void SyncPageNames()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;
            MSVisio.Document doc = app.ActiveDocument;
            MSVisio.Page page = app.ActivePage;

            SyncPageNames(page);
        }

        //    'Restore diagram services
        //    ActiveDocument.DiagramServicesEnabled = DiagramServices
        public static void SyncPageNames(MSVisio.Page page)
        {
            Common.WriteToDebugWindow(string.Format("{0} Name:>{1}< NameU:>{2}<",
                MethodBase.GetCurrentMethod().Name, page.Name, page.NameU));

            try
            {
                Common.VisioApplication.EventsEnabled = 0;
                page.NameU = page.Name;
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force:true);

            }
            finally
            {
                Common.VisioApplication.EventsEnabled = 1;
            }
        }

        //    Application.ActivePage.Name = "QMS RDL <DataSources>"
        public static void ToggleLayerLock(MSVisio.Application app, string doc, string page, string shape, string shapeu)
        {
            Common.WriteToDebugWindow(string.Format("{0})",
                MethodBase.GetCurrentMethod().Name));

            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
            string layerName = activeShape.CellsU["Prop.Layer"].ResultStrU[0];

            foreach (MSVisio.Layer layer in app.ActivePage.Layers)
            {

                Common.WriteToDebugWindow(layer.Name);

                if (layer.Name.ToLower() == layerName.ToLower())
                {
                    var currentState = layer.CellsC[(short)MSVisio.VisCellIndices.visLayerLock].ResultIU;
                    string newState = null;

                    newState = (currentState == 0) ? "1" : "0";

                    //bool state = !bool.Parse(.ToString());
                    layer.CellsC[(short)MSVisio.VisCellIndices.visLayerLock].Formula = newState;
                    activeShape.CellsU["Prop.Lock"].FormulaU = newState;
                }
            }
        }

        //    Application.ActiveWindow.Page.Paste
        public static void ToggleLayerPrint(MSVisio.Application app, string doc, string page, string shape, string shapeu)
        {
            Common.WriteToDebugWindow(string.Format("{0})",
                MethodBase.GetCurrentMethod().Name));

            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
            string layerName = activeShape.CellsU["Prop.Layer"].ResultStrU[0];

            foreach (MSVisio.Layer layer in app.ActivePage.Layers)
            {

                Common.WriteToDebugWindow(layer.Name);

                if (layer.Name.ToLower() == layerName.ToLower())
                {
                    var currentState = layer.CellsC[(short)MSVisio.VisCellIndices.visLayerPrint].ResultIU;
                    string newState = null;

                    newState = (currentState == 0) ? "1" : "0";

                    //bool state = !bool.Parse(.ToString());
                    layer.CellsC[(short)MSVisio.VisCellIndices.visLayerPrint].Formula = newState;
                    activeShape.CellsU["Prop.Print"].FormulaU = newState;
                }
            }
        }

        public static void ToggleLayerSetting(MSVisio.Page activePage, MSVisio.Shape activeShape, MSVisio.VisCellIndices visCell)
        {
            string layerName = activeShape.CellsU["Prop.Layer"].ResultStrU[0];

            foreach (MSVisio.Layer layer in activePage.Layers)
            {
                Common.WriteToDebugWindow(layer.Name);

                if (layer.Name.ToLower() == layerName.ToLower())
                {
                    var currentState = layer.CellsC[(short)visCell].ResultIU;
                    string newState = null;

                    newState = (currentState == 0) ? "1" : "0";

                    //bool state = !bool.Parse(.ToString());
                    layer.CellsC[(short)visCell].Formula = newState;
                    activeShape.CellsU["Prop.Visible"].FormulaU = newState;
                }
            }
        }

        //    Dim UndoScopeID1 As Long
        //    UndoScopeID1 = Application.BeginUndoScope("Insert Page")
        //    Dim vsoPage1 As MSVisio.Page
        //    Set vsoPage1 = ActiveDocument.Pages.Add
        //    vsoPage1.Name = "Page-54"
        //    vsoPage1.Background = False
        //    vsoPage1.Index = 54
        //    vsoPage1.BackPage = "Page Base"
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageWidth).FormulaU = "11 in"
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageHeight).FormulaU = "8.5 in"
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPage, visPageDrawSizeType).FormulaU = "3"
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPage, 38).FormulaU = "2"
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPageLayout, visPLOPlaceStyle).FormulaForceU = "0"
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPageLayout, visPLORouteStyle).FormulaForceU = "0"
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPageLayout, visPLOPlowCode).FormulaForceU = "0"
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPageLayout, visPLOJumpCode).FormulaForceU = "1"
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPageLayout, visPLOJumpStyle).FormulaForceU = "0"
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPageLayout, visPLOLineAdjustFrom).FormulaForceU = "0"
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPageLayout, visPLOLineAdjustTo).FormulaForceU = "0"
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPageLayout, visPLOLineRouteExt).FormulaForceU = "0"
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPageLayout, visPLOSplit).FormulaForceU = "1"
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowPrintProperties, visPrintPropertiesPageOrientation).FormulaU = "2"
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowThemeProperties, visColorSchemeIndex).FormulaU = ""
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowThemeProperties, visEffectSchemeIndex).FormulaU = ""
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowThemeProperties, visConnectorSchemeIndex).FormulaU = ""
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowThemeProperties, visFontSchemeIndex).FormulaU = ""
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowThemeProperties, visThemeIndex).FormulaU = ""
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowThemeProperties, visVariationColorIndex).FormulaU = ""
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowThemeProperties, visVariationStyleIndex).FormulaU = ""
        //    vsoPage1.PageSheet.CellsSRC(visSectionObject, visRowThemeProperties, visEmbellishmentIndex).FormulaU = ""
        //    Application.EndUndoScope UndoScopeID1, True
        public static void ToggleLayerVisibility(MSVisio.Application app, string doc, string page, string shape, string shapeu)
        {
            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name}");

            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
            MSVisio.Page activePage = app.ActivePage;

            ToggleLayerSetting(activePage, activeShape, MSVisio.VisCellIndices.visLayerVisible);
            //string layerName = activeShape.CellsU["Prop.Layer"].ResultStrU[0];

            //foreach (MSVisio.Layer layer in activePage.Layers)
            //{

            //    Common.WriteToDebugWindow(layer.Name);

            //    if (layer.Name.ToLower() == layerName.ToLower())
            //    {
            //        var currentState = layer.CellsC[(short)MSVisio.VisCellIndices.visLayerVisible].ResultIU;
            //        string newState = null;

            //        newState = (currentState == 0) ? "1" : "0";

            //        //bool state = !bool.Parse(.ToString());
            //        layer.CellsC[(short)MSVisio.VisCellIndices.visLayerVisible].Formula = newState;
            //        activeShape.CellsU["Prop.Visible"].FormulaU = newState;
            //    }
            //}
        }

        public static void UnlockLayer(string layerName)
        {
            Common.WriteToDebugWindow(string.Format("{0}()",
                MethodBase.GetCurrentMethod().Name));

            MSVisio.Layer layer = Common.VisioApplication.ActivePage.Layers[layerName];

            layer.CellsC[(short)MSVisio.VisCellIndices.visLayerLock].Formula = "0";
        }
        //Sub MovePageTryTwo()

        //    'Enable diagram services
        //    Dim DiagramServices As Integer
        //    DiagramServices = ActiveDocument.DiagramServicesEnabled
        //    ActiveDocument.DiagramServicesEnabled = visServiceVersion140 + visServiceVersion150

        //    Application.ActiveWindow.Page = Application.ActiveDocument.Pages.ItemU("QMS RDL <DataSources>")

        //    Application.ActiveWindow.SelectAll

        //    ActiveWindow.DeselectAll
        //    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(2), visSelect
        //    ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromID(1), visSelect
        //    Application.ActiveWindow.Selection.Cut
        //    Application.Windows.ItemEx("CHR Notes-BD - QMS Reports.vsdx").Activate
        // TODO(crhodes):
        // This method has become a mess.  It supports two versions of Color Pickers and has been altered by use of a "mode" argument.
        // May want to have two routines to keep the code simpler.  For now leave as is.
        // Main difference is in what comes in the propColorS and UserColorS variables.
        // 
        // No mode passed, e.g. UpdateGroupNameShapes
        //  User picks a color by name and an index looks up the RGB value
        //
        // Mode passed and = 1, e.g. UpdateGroupNameShapes,1
        //  User has updated a RGB value, e.g. RGB(10,20,30)

        public static void UpdateGroupNameShapes(MSVisio.Application app, string doc, string page, string shape, string shapeu, String[] args)
        {
            string mode = null;

            if (args.Count() != 1)
            {
                //Common.WriteToDebugWindow(string.Format("Incorrect Argument Count, expected 3.  Check ShapeSheet"));
            }
            else
            {
                mode = args[0];
            }

            Common.WriteToDebugWindow(string.Format("{0}({1})  mode:{2}",
                MethodBase.GetCurrentMethod().Name, page,
                mode));

            MSVisio.Page currentPage = app.ActivePage;
            MSVisio.Shape colorSelectorShape = currentPage.Shapes[shape];
            string colorSelectorGroupName = null;

            if (colorSelectorShape.CellExists["Prop.GroupName", 0] != 0)
            {
                colorSelectorGroupName = colorSelectorShape.Cells["Prop.GroupName"].ResultStrU[0];
            }
            else
            {
                Common.WriteToDebugWindow(string.Format("Cannot locate Prop.GroupName.  Check ShapeSheet"));
                return;
            }

            MSVisio.Cell userCell;
            MSVisio.Cell propCell;
            MSVisio.Cell propCell2;
            double userColor = double.NaN;
            double propColor = double.NaN;
            double propColor2 = double.NaN;

            string userColorS = null;
            string propColorS = null;
            string propColor2S = null;

            // Extract color information from the Color Selector Tool (Current Shape)
            if (null == mode)
            {
                userCell = colorSelectorShape.Cells["User.Color"];
                propCell = colorSelectorShape.Cells["Prop.Color"];

                userColor = userCell.ResultIU;
                propColor = propCell.ResultIU;

                userColorS = userCell.ResultStrU[0];
                //propColorS = propCell.ResultStrU[0];
                propColorS = userColorS;
            }
            else
            {
                propCell = colorSelectorShape.Cells["Prop.Color"];
                propColorS = propCell.ResultStrU[0];
                userColorS = propColorS;

                // Some selector tools support more than one color selector

                if (colorSelectorShape.CellExistsU["Prop.Color2", 0] != 0)
                {
                    propColor2S = colorSelectorShape.CellsU["Prop.Color2"].ResultStrU[0];
                }
            }

            Common.WriteToDebugWindow(string.Format("userColor:{0}-{1} propColor:{2}-{3} {4}", userColor, userColorS, propColor, propColorS, propColor2S));

            // Now walk the shapes on the page looking for shapes with a matching GroupName

            string groupName = string.Empty;
            short isSelectorTool = 0;
            short hasGroupName = 0;

            foreach (MSVisio.Shape pageShape in currentPage.Shapes)
            {
                Common.WriteToDebugWindow(string.Format("shape NameID:({0})", pageShape.NameID));
                try
                {
                    if ((hasGroupName = pageShape.CellExistsU["Prop.GroupName", 0]) != 0)
                    {
                        groupName = pageShape.CellsU["Prop.GroupName"].ResultStrU[0];
                    }
                    else
                    {
                        groupName = "";
                    }

                    if (hasGroupName != 0)
                    {
                        if (colorSelectorGroupName.Equals(groupName))
                        {
                            //    var isSelectorTool = shape.CellExists["User.IsPageName", 0]; // 0 is Local and Inherited, 1 is Local only 
                            isSelectorTool = pageShape.CellExistsU["User.IsSelectorTool", 0];

                            // Not all Shapes with GroupName have the isSelectorTool user property, e.g. the Color Selector Shape!
                            // We don't care what the value is, we only update the shapes that don't have the user property.

                            Common.WriteToDebugWindow(string.Format("   groupName:({0})  isSelectorTool:({1})", groupName, isSelectorTool));

                            if (isSelectorTool == 0)
                            {
                                pageShape.CellsU["Prop.Color"].FormulaU = string.Format("\"{0}\"", propColorS);

                                if (propColor2S != null)
                                {
                                    // See if the shape supports a second color

                                    if (pageShape.CellExistsU["Prop.Color2", 0] != 0)
                                    {
                                        pageShape.CellsU["Prop.Color2"].FormulaU = string.Format("\"{0}\"", propColor2S);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow(ex.ToString(), force:true);
                }
            }
        }

        public static void UpdateHasColorTagsShapes(MSVisio.Application app, string doc, string page, string shape, string shapeu, String[] args)
        {
            int levels = 0;

            if (args.Count() != 1)
            {
                //Common.WriteToDebugWindow(string.Format("Incorrect Argument Count, expected 3.  Check ShapeSheet"));
            }
            else
            {
                levels = int.Parse(args[0]);
            }

            Common.WriteToDebugWindow(string.Format("{0}({1})  levels:{2}",
                MethodBase.GetCurrentMethod().Name, page,
                levels));

            MSVisio.Page currentPage = app.ActivePage;
            MSVisio.Shape colorTagShape = currentPage.Shapes[shape];

            string tagName = null;
            string foregroundColor = null;
            string backgroundColor = null;
            string pattern = null;
            string isVisible = null;

            if (colorTagShape.CellExists["Prop.TagName", 0] != 0)
            {
                tagName = colorTagShape.Cells["Prop.TagName"].ResultStrU[0];
            }
            else
            {
                Common.WriteToDebugWindow(string.Format("Cannot locate Prop.TagName.  Check ShapeSheet"));
                return;
            }

            if (colorTagShape.CellExists["Prop.ForegroundColor", 0] != 0)
            {
                foregroundColor = colorTagShape.Cells["Prop.ForegroundColor"].ResultStrU[0];
            }
            else
            {
                Common.WriteToDebugWindow(string.Format("Cannot locate Prop.ForegroundColor.  Check ShapeSheet"));
                return;
            }

            if (colorTagShape.CellExists["Prop.BackgroundColor", 0] != 0)
            {
                backgroundColor = colorTagShape.Cells["Prop.BackgroundColor"].ResultStrU[0];
            }
            else
            {
                Common.WriteToDebugWindow(string.Format("Cannot locate Prop.BackgroundColor.  Check ShapeSheet"));
                return;
            }

            if (colorTagShape.CellExists["Prop.Pattern", 0] != 0)
            {
                pattern = colorTagShape.Cells["Prop.Pattern"].ResultStrU[0];
            }
            else
            {
                Common.WriteToDebugWindow(string.Format("Cannot locate Prop.Pattern.  Check ShapeSheet"));
                return;
            }

            if (colorTagShape.CellExists["Prop.IsVisible", 0] != 0)
            {
                isVisible = colorTagShape.Cells["Prop.IsVisible"].ResultStrU[0];
            }
            else
            {
                Common.WriteToDebugWindow(string.Format("Cannot locate Prop.IsVisible.  Check ShapeSheet"));
                return;
            }

            foreach (MSVisio.Shape pageShape in currentPage.Shapes)
            {
                Common.WriteToDebugWindow(string.Format("shape NameID:({0})", pageShape.NameID));

                try
                {
                    var hasColorTags = pageShape.CellExistsU["User.HasColorTags", 0];    // 0 is Local and Inherited, 1 is Local only 

                    Common.WriteToDebugWindow(string.Format("shape {0}  hasColorTags:{1})",
                        pageShape.Name, hasColorTags));

                    if (hasColorTags != 0)
                    {
                        foreach (MSVisio.Shape subShape in pageShape.Shapes)
                        {
                            if (subShape.CellExistsU["Prop.TagName", 0] != 0)
                            {
                                if (tagName == subShape.CellsU["Prop.TagName"].ResultStrU[0])
                                {
                                    if (subShape.CellExistsU["Prop.ForegroundColor", 0] != 0)
                                    {
                                        subShape.CellsU["Prop.ForegroundColor"].FormulaU = foregroundColor.WrapInDblQuotes();
                                    }

                                    if (subShape.CellExistsU["Prop.BackgroundColor", 0] != 0)
                                    {
                                        subShape.CellsU["Prop.BackgroundColor"].FormulaU = backgroundColor.WrapInDblQuotes();
                                    }

                                    if (subShape.CellExistsU["Prop.Pattern", 0] != 0)
                                    {
                                        subShape.CellsU["Prop.Pattern"].FormulaU = pattern;
                                    }

                                    if (subShape.CellExistsU["Prop.IsVisible", 0] != 0)
                                    {
                                        subShape.CellsU["Prop.IsVisible"].FormulaU = isVisible;
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow(ex.ToString(), force:true);
                }
            }
        }

        public static void UpdateLayer(MSVisio.Application app, string doc, string page, string shape, string shapeu)
        {
            Common.WriteToDebugWindow(string.Format("{0})",
                MethodBase.GetCurrentMethod().Name));

            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
            string layerName = activeShape.CellsU["Prop.Layer"].ResultStrU[0];

            foreach (MSVisio.Layer layer in app.ActivePage.Layers)
            {
                try
                {
                    Common.WriteToDebugWindow(layer.Name);

                    if (layer.Name.ToLower() == layerName.ToLower())
                    {
                        if (activeShape.CellExistsU["Prop.Visible", 0] != 0)
                        {
                            layer.CellsC[(short)MSVisio.VisCellIndices.visLayerVisible].FormulaU = activeShape.CellsU["Prop.Visible"].ResultStrU[0];
                        }

                        if (activeShape.CellExistsU["Prop.Lock", 0] != 0)
                        {
                            layer.CellsC[(short)MSVisio.VisCellIndices.visLayerLock].FormulaU = activeShape.CellsU["Prop.Lock"].ResultStrU[0];
                        }

                        if (activeShape.CellExistsU["Prop.Print", 0] != 0)
                        {
                            layer.CellsC[(short)MSVisio.VisCellIndices.visLayerPrint].FormulaU = activeShape.CellsU["Prop.Print"].ResultStrU[0];
                        }
                    }
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow(ex.ToString(), force:true);
                }
            }
        }

        public static void UpdatePageNameShapes(MSVisio.Page page)
        {
            Common.WriteToDebugWindow(string.Format("{0} Name:>{1}< NameU:>{2}<",
                MethodBase.GetCurrentMethod().Name, page.Name, page.NameU));

            foreach (MSVisio.Shape shape in page.Shapes)
            {
                Actions.Visio_Shape.UpdatePageNameShape(shape, page.Name);
            }
        }

        #endregion

        #region Private Methods

        private static List<MSVisio.Shape> GetLayerNameShapes(MSVisio.Page page, LayerNameType nameType = LayerNameType.AllNames)
        {
            Common.WriteToDebugWindow(string.Format("{0}({1})",
                MethodBase.GetCurrentMethod().Name, page.NameU));

            List<MSVisio.Shape> layerNames = new List<MSVisio.Shape>();

            foreach (MSVisio.Shape shape in page.Shapes)
            {
                if (shape.CellExistsU["User.IsLayerName", 0] != 0)
                {
                    switch (nameType)
                    {
                        case LayerNameType.AllNames:
                            layerNames.Add(shape);

                            break;

                        case LayerNameType.AddName:
                            if (shape.CellExistsU["User.AddName", 0] != 0)
                            {
                                layerNames.Add(shape);
                            }

                            break;

                        case LayerNameType.RemovalName:
                            if (shape.CellExistsU["User.RemovalName", 0] != 0)
                            {
                                layerNames.Add(shape);
                            }

                            break;

                        default:
                            Common.WriteToDebugWindow(string.Format("Unknown LayerNameype:{0}",
                                nameType));
                            break;
                    }
                }
            }

            return layerNames;
        }

        private static List<MSVisio.Shape> GetNavigationLinks(MSVisio.Page page)
        {
            Common.WriteToDebugWindow(string.Format("{0}({1})",
                MethodBase.GetCurrentMethod().Name, page.NameU));

            List<MSVisio.Shape> navigationLinks = new List<MSVisio.Shape>();

            foreach (MSVisio.Shape shape in page.Shapes)
            {
                var isNavigationLink = shape.CellExists["User.IsNavigationLink", 0];

                navigationLinks.Add(shape);
            }

            return navigationLinks;
        }

        private static string GetPageSaveName(MSVisio.Page page)
        {
            string pageName = VNCVisioAddIn.Helpers.SafePageName(page.NameU);
            string documentName = VNCVisioAddIn.Helpers.SafeFileName(page.Application.ActiveDocument.Name);

            // TODO(crhodes):
            // Do more fancy stuff so it is easier to find the file later

            pageName = string.Format(@"C:\temp\VisioExport\{0}-{1}.png", documentName, pageName);

            return pageName;
        }

        private static void RemoveNavigationLinks(MSVisio.Page page)
        {
            Common.WriteToDebugWindow(string.Format("{0} Name:>{1}< NameU:>{2}<",
                MethodBase.GetCurrentMethod().Name, page.Name, page.NameU));

            List<MSVisio.Shape> navigationLinks = GetNavigationLinks(page);

            try
            {
                foreach (MSVisio.Shape shape in navigationLinks)
                {
                    var isNavigationLink = shape.CellExists["User.IsNavigationLink", 0];  // 0 not limited to local only

                    if (isNavigationLink != 0)
                    {
                        shape.Delete();
                    }
                }
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force:true);
            }
        }

        #endregion

    }
}
