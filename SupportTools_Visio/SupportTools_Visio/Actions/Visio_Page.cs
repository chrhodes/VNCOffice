using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Visio = Microsoft.Office.Interop.Visio;
using VisioHelper = VNC.AddinHelper.Visio;
using AIH = VNC.AddinHelper;
using VNC;
using VNC.Core;
using System.Text.RegularExpressions;
using DevExpress.XtraRichEdit.Unicode.TextAnalyzer;
using Microsoft.Office.Interop.Visio;

namespace SupportTools_Visio.Actions
{
    class Visio_Page
    {
        #region Enums, Fields, Properties, Structures

        private enum LayerNameType
        {
            AllNames = 0,
            AddName = 1,
            RemovalName = 2
        }

        #endregion

        #region Main Methods

        public static void AddDefaultLayers()
        {
            VNC.Log.Trace("", Common.LOG_CATEGORY, 0);

            VisioHelper.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Application app = Globals.ThisAddIn.Application;
            Visio.Document doc = app.ActiveDocument;
            Visio.Page page = app.ActivePage;

            AddDefaultLayers(page);
        }

        public static Visio.Layer AddLayer(Visio.Page page, 
            string layerName,
            string layerVisible = "1", 
            string layerPrint = "1", 
            string layerActive = "0", 
            string layerLock = "0", 
            string layerSnap = "1", 
            string layerGlue = "1")
        {
            Visio.Layer layer = null;

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

                layer.CellsC[(short)Visio.VisCellIndices.visLayerVisible].FormulaU = layerVisible;

                layer.CellsC[(short)Visio.VisCellIndices.visLayerPrint].FormulaU = layerPrint;

                layer.CellsC[(short)Visio.VisCellIndices.visLayerActive].FormulaU = layerActive;

                layer.CellsC[(short)Visio.VisCellIndices.visLayerLock].FormulaU = layerLock;

                layer.CellsC[(short)Visio.VisCellIndices.visLayerSnap].FormulaU = layerSnap;

                layer.CellsC[(short)Visio.VisCellIndices.visLayerGlue].FormulaU = layerGlue;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }

            return layer;
        }

        public static void AutoSizePageOn()
        {
            VNC.Log.Trace("", Common.LOG_CATEGORY, 0);

            VisioHelper.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Application app = Globals.ThisAddIn.Application;
            Visio.Document doc = app.ActiveDocument;
            Visio.Page page = app.ActivePage;

            AutoSizePageOn(page);
        }

        public static void AutoSizePageOff()
        {
            VNC.Log.Trace("", Common.LOG_CATEGORY, 0);

            VisioHelper.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Application app = Globals.ThisAddIn.Application;
            Visio.Document doc = app.ActiveDocument;
            Visio.Page page = app.ActivePage;

            AutoSizePageOff(page);
        }

        public static void AutoSizePageOn(Visio.Page page)
        {
            page.AutoSize = true;
        }

        public static void AutoSizePageOff(Visio.Page page)
        {
            page.AutoSize = false;
        }

        public static void LockLayer(string layerName)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Layer layer = Globals.ThisAddIn.Application.ActivePage.Layers[layerName];

            layer.CellsC[(short)Visio.VisCellIndices.visLayerLock].Formula = "1";
        }

        public static void UnlockLayer(string layerName)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Layer layer = Globals.ThisAddIn.Application.ActivePage.Layers[layerName];

            layer.CellsC[(short)Visio.VisCellIndices.visLayerLock].Formula = "0";
        }

        public static void AddDefaultLayers(Visio.Page page)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0}({1})",
                System.Reflection.MethodInfo.GetCurrentMethod().Name,
                page.NameU));

            if (page == null)
            {
                System.Windows.Forms.MessageBox.Show("No ActivePage");
                return;
            }

            Visio.Layers layers = page.Layers;

            try
            {
                Visio.Page layersPage = Globals.ThisAddIn.Application.ActiveDocument.Pages["Default Layers"];
                //VisioHelper.DisplayInWatchWindow(string.Format("  Copying {0} links", linkPage.Shapes.Count));

                List<Visio.Shape> layerNames = GetLayerNameShapes(layersPage, LayerNameType.AddName);
            
                string layerName = null;

                // These are the defaults if the shape does not have values.

                string layerVisible = "1";
                string layerPrint = "1";
                string layerActive = "0";
                string layerLock = "0";
                string layerSnap = "1";
                string layerGlue = "1";
                
                foreach (Visio.Shape shape in layerNames)
                {
                    AddLayer(page, shape.Text, layerVisible, layerPrint, layerActive, layerLock, layerSnap, layerGlue);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
                // No navigation Links Page perhaps
            }
        }

        public static void AddNavigationLinks(Visio.Page page)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0}({1})",
                System.Reflection.MethodInfo.GetCurrentMethod().Name, 
                page.NameU));

            if ((page.Background != 0) || (page.NameU == "Navigation Links"))
            {
                //VisioHelper.DisplayInWatchWindow("   Skipping");
            	return;
            }

            RemoveNavigationLinks(page);

            //List<Visio.Shape> links = Actions.Visio_Document.GetNavigationLinks();

            try
            {
                Visio.Window activeWindow = Globals.ThisAddIn.Application.ActiveWindow;
                activeWindow.Page = Globals.ThisAddIn.Application.ActiveDocument.Pages["Navigation Links"];
                activeWindow.SelectAll();
                activeWindow.Selection.Copy(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate);
                activeWindow.Page = page;
                page.Paste(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate);

                //Globals.ThisAddIn.Application.Windows.ItemEx["Navigation Links"].Activate();
                //Globals.ThisAddIn.Application.ActiveWindow.SelectAll();
                //Globals.ThisAddIn.Application.ActiveWindow.Selection.Copy();
                //Globals.ThisAddIn.Application.Windows.ItemEx["Navigation Links"].Activate();


                //Visio.Page linkPage = Globals.ThisAddIn.Application.ActiveDocument.Pages["Navigation Links"];
                //linkPage.Application.
                //Globals.ThisAddIn.Application.
                //VisioHelper.DisplayInWatchWindow(string.Format("  Copying {0} links", linkPage.Shapes.Count));

                //foreach (Visio.Shape shape in linkPage.Shapes)
                //{
                //    // TODO: Make this smarter about only using IsNavigationLink shapes
                //    shape.Copy(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate);
                //    page.Paste(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate);
                //}

                // Typically we don't print the stuff on the navigation layer.

                page.Layers["Navigation"].CellsC[(short)Visio.VisCellIndices.visLayerPrint].FormulaU = "0";
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
                // No navigation Links Page perhaps
            }
        }

        public static void CreateDefaultLayersPage(Visio.Application app, string doc, string page, string shape, string shapeu, string[] array)
        {
            Visio.Page newPage = CreatePage(pageName: "Default Layers", backgroundPageName: "", isBackground: 1);
        }

        public static void CreateNavigationLinksPage(Visio.Application app, string doc, string page, string shape, string shapeu, string[] array)
        {
            Visio.Page newPage = CreatePage(pageName: "Navigation Links", backgroundPageName: "", isBackground: 1);
        }

        public static void CreatePageBasePage(Visio.Application app, string doc, string page, string shape, string shapeu, string[] array)
        {
            Visio.Page newPage = CreatePage(pageName: "Page Base", backgroundPageName: "", isBackground: 1);
        }

        public static Visio.Page CreatePage(string pageName, string backgroundPageName, short isBackground = 0)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0}() Page:{1}  Background:{2}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name,
                pageName, backgroundPageName));

            // TODO(crhodes):
            //	Error handling. Page already exists, background page doesn't exist, etc.
            Visio.Application app = Globals.ThisAddIn.Application;
            int currentPageIndex = app.ActivePage.Index;

            VisioHelper.DisplayInWatchWindow(string.Format("   currentPageIndex:{0}", currentPageIndex));
            Visio.Page newPage = app.ActiveDocument.Pages.Add();

            // Cleanup page names
            pageName = pageName.Replace("\n", " ");

            newPage.Name = pageName;

            try
            {
                newPage.BackPage = backgroundPageName;
            }
            catch (Exception ex)
            {
                VisioHelper.DisplayInWatchWindow(string.Format("Cannot Find Background Page ({0})", backgroundPageName));
            }
            
            newPage.Index = (short)(currentPageIndex + 1);

            newPage.Background = isBackground;

            AddNavigationLinks(newPage);

            return newPage;
        }

        public static void CreateActivityPage(Visio.Application app, string doc, string page, string shape, string shapeu, string[] args)
        {
            string pageLevel = null;
            string backgroundPageName = null;

            if (args.Count() != 2)
            {
                VisioHelper.DisplayInWatchWindow(string.Format("Incorrect Argument Count, expected 2.  Check ShapeSheet"));
            }
            else
            {
                pageLevel = args[0];
                backgroundPageName = args[1];
            }

            VisioHelper.DisplayInWatchWindow(string.Format("{0}() PageLevel:{1}  Background:{2}", 
                System.Reflection.MethodInfo.GetCurrentMethod().Name,
                pageLevel, backgroundPageName));

            // Current shape contains text for new page name.

            Visio.Shape activeShape = app.ActivePage.Shapes[shape];
            VisioHelper.DisplayInWatchWindow(string.Format("  Shape(Name:{0}  Text:{1}", activeShape.Name, activeShape.Text));

            //string newPageName = string.Format("{0}-{1}", pageLevel, activeShape.Text);
            string newPageName = string.Format("{0}{1}{2}", pageLevel, "-", activeShape.Text);

            Visio.Page newPage = CreatePage(newPageName, backgroundPageName);

            // Update the current shape's hyperlink to point to the new page

            // TODO(crhodes):
            //	Not sure which of these two approaches is doing the magic.

            Visio.Hyperlink currentHyperLink = activeShape.AddHyperlink();
            currentHyperLink.SubAddress = newPageName;

            //activeShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionHyperlink, 0, 2].Formula = newPageName;


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

        public static void CreateArtifactPage(Visio.Application app, string doc, string page, string shape, string shapeu, string[] args)
        {
            string pageLevel = null;
            string backgroundPageName = null;

            if (args.Count() != 2)
            {
                VisioHelper.DisplayInWatchWindow(string.Format("Incorrect Argument Count, expected 2.  Check ShapeSheet"));
            }
            else
            {
                pageLevel = args[0];
                backgroundPageName = args[1];
            }

            VisioHelper.DisplayInWatchWindow(string.Format("{0}() PageLevel:{1}  Background:{2}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name,
                pageLevel, backgroundPageName));

            // Current shape contains text for new page name.

            Visio.Shape activeShape = app.ActivePage.Shapes[shape];
            VisioHelper.DisplayInWatchWindow(string.Format("  Shape(Name:{0}  Text:{1}", activeShape.Name, activeShape.Text));

            //string newPageName = string.Format("{0}-{1}", pageLevel, activeShape.Text);
            string newPageName = string.Format("{0}{1}{2}", "", "", activeShape.Text);

            Visio.Page newPage = CreatePage(newPageName, backgroundPageName);

            // Update the current shape's hyperlink to point to the new page

            // TODO(crhodes):
            //	Not sure which of these two approaches is doing the magic.

            Visio.Hyperlink currentHyperLink = activeShape.AddHyperlink();
            currentHyperLink.SubAddress = newPageName;

            activeShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionHyperlink, 0, 2].Formula = newPageName;

            //  Figure out how to get PageName shape added.
        }

        public static void CreateMetricPage(Visio.Application app, string doc, string page, string shape, string shapeu, String[] args)
        {
            string arg0 = null;
            string backgroundPageName = null;

            if (args.Count() != 2)
            {
                VisioHelper.DisplayInWatchWindow(string.Format("Incorrect Argument Count, expected 2.  Check ShapeSheet"));
            }
            else
            {
                arg0 = args[0];
                backgroundPageName = args[1];
            }

            VisioHelper.DisplayInWatchWindow(string.Format("{0}() arg0:{1}  Background:{2}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name,
                arg0, backgroundPageName));

            // Current shape contains text for new page name.

            Visio.Shape activeShape = app.ActivePage.Shapes[shape];
            VisioHelper.DisplayInWatchWindow(string.Format("  Shape(Name:{0}  Text:{1}", activeShape.Name, activeShape.Text));

            string newPageName = string.Format("{0}{1}{2}", "", "", activeShape.Text);

            Visio.Page newPage = CreatePage(newPageName, backgroundPageName);

            Visio.Hyperlink currentHyperLink = activeShape.AddHyperlink();
            currentHyperLink.SubAddress = newPageName;

            activeShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionHyperlink, 0, 2].Formula = newPageName;

            // TODO(crhodes):
            //	Figure out what to do with roleSource
            //  Figure out how to get PageName shape added.
        }

        public static void CreatePageForShape(Visio.Application app, string doc, string page, string shape, string shapeu, string[] args)
        {
            string prefix = null;
            string delimiter = null;
            string backgroundPageName = null;

            if (args.Count() != 3)
            {
                VisioHelper.DisplayInWatchWindow(string.Format("Incorrect Argument Count, expected 3.  Check ShapeSheet"));
            }
            else
            {
                prefix = args[0];
                delimiter = args[1];
                backgroundPageName = args[2];
            }

            VisioHelper.DisplayInWatchWindow(string.Format("{0}() prefix:>{1}< delimiter:>{2}< backgroundPageName:>{3}<",
                System.Reflection.MethodInfo.GetCurrentMethod().Name,
                prefix, delimiter, backgroundPageName));

            try
            {
                // Current shape contains text for new page name.
                Visio.Page activePage = app.ActivePage;
                Visio.Shape activeShape = app.ActivePage.Shapes[shape];
                VisioHelper.DisplayInWatchWindow(string.Format("  Shape(Name:>{0}< Text:>{1}< Characters:>{2}<", activeShape.Name, activeShape.Text, activeShape.Characters.TextAsString));

                string shapePageName = "Error-PageNameNotProvided";

                if (activeShape.CellExistsU["Prop.PageName", 0] != 0)
                {
                    shapePageName = activeShape.CellsU["Prop.PageName"].ResultStrU[Visio.VisUnitCodes.visUnitsString];
                }
                else if (activeShape.Characters.TextAsString.Length > 0)
                {
                    //string newPageName = string.Format("{0}{1}{2}", prefix, delimiter, activeShape.Text);
                    // shape.Text comes in as OBJ if use fields and Shape Data.   Use shape.Characters instead.
                    shapePageName = activeShape.Characters.TextAsString;
                }
              
                string newPageName = string.Format("{0}{1}{2}", prefix, delimiter, shapePageName);
                
                Visio.Page newPage = CreatePage(newPageName, backgroundPageName);

                // The old style linkable masters did not have Prop.Data for the HyperLink.  Check first before updating.
                // Should really retire all the old shapes and remove this code.

                if (activeShape.CellExistsU["Prop.HyperLink", 0] != 0)
                {
                    activeShape.CellsU["Prop.HyperLink"].FormulaU = newPageName.WrapInDblQuotes();
                }
                else
                {
                    Visio.Hyperlink currentHyperLink = activeShape.AddHyperlink();
                    currentHyperLink.SubAddress = newPageName;
                }

                // Check to see if there is a ReturnLink Property with values that can be used to create a return link
                // to the page that linked to us.

                if (activeShape.CellExistsU["Prop.ReturnLink", 0] != 0)
                {
                    //string returnLinkProp = activeShape.CellsU["Prop.ReturnLink"].FormulaU;   // This returns "<string>"  we want just <string>
                    string returnLinkProp = activeShape.CellsU["Prop.ReturnLink"].ResultStrU[Visio.VisUnitCodes.visUnitsString];
                    string[] linkInfo = returnLinkProp.Split(',');
                    string stencilName = linkInfo[0];
                    string shapeName = linkInfo[1];

                    VisioHelper.DisplayInWatchWindow(string.Format("  returnLinkProp:>{0}< stencilName:>{1}< shapeName:>{2}< ", returnLinkProp, stencilName, shapeName));

                    try
                    {
                        Visio.Document linkStencil = app.Documents[stencilName];

                        try
                        {
                             Visio.Master linkMaster = linkStencil.Masters[shapeName];

                            // Add return link in upper left corner.  Assume 11x8.5

                            // TODO(crhodes)
                            // Get Page Size and drop in upper left
                            Visio.Shape returnLinkShape = newPage.Drop(linkMaster, 1.0, 8.0);

                            returnLinkShape.CellsU["Prop.PageName"].FormulaU = activePage.Name.WrapInDblQuotes();
                            returnLinkShape.CellsU["Prop.HyperLink"].FormulaU = activePage.Name.WrapInDblQuotes();
                        }
                        catch (Exception ex)
                        {
                            VisioHelper.DisplayInWatchWindow(string.Format("  Cannot find Master named:>{0}<", shapeName));
                        }
                    }
                    catch (Exception ex)
                    {
                        VisioHelper.DisplayInWatchWindow(string.Format("  Cannot find open Stencil named:>{0}<", stencilName));
                    }
                }

                // Add a header.  May want to pick the stencil and shape for config file.
                // Or add a property to Shape.

                VisioHelper.LoadStencil(app, "Page Shapes.vssx");
                Visio.Master headerMaster = app.Documents[@"Page Shapes.vssx"].Masters[@"18pt Header"];

                newPage.Drop(headerMaster, 5.5, 8.0625);

                // NOTE(crhodes)
                // Add the shape that triggered the event.  User can delete if doesn't want.
                // More and more often I go back and copy it, traverse the link, and drop it.
                // Drop in middle of page for now assuming 11x8.5

                // TODO(crhodes)
                // Get Page Size and drop in center
                newPage.Drop(activeShape,5.5,4.25);
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void CreateRolePage(Visio.Application app, string doc, string page, string shape, string shapeu, String[] args)
        {
            string roleSource = null;
            string backgroundPageName = null;

            if (args.Count() != 2)
            {
                VisioHelper.DisplayInWatchWindow(string.Format("Incorrect Argument Count, expected 2.  Check ShapeSheet"));
            }
            else
            {
                roleSource = args[0];
                backgroundPageName = args[1];
            }

            VisioHelper.DisplayInWatchWindow(string.Format("{0}() PageLevel:{1}  Background:{2}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name,
                roleSource, backgroundPageName));

            // Current shape contains text for new page name.

            Visio.Shape activeShape = app.ActivePage.Shapes[shape];
            VisioHelper.DisplayInWatchWindow(string.Format("  Shape(Name:{0}  Text:{1}", activeShape.Name, activeShape.Text));

            string newPageName = string.Format("{0}{1}{2}", "", "", activeShape.Text);

            Visio.Page newPage = CreatePage(newPageName, backgroundPageName);

            Visio.Hyperlink currentHyperLink = activeShape.AddHyperlink();
            currentHyperLink.SubAddress = newPageName;

            activeShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionHyperlink, 0, 2].Formula = newPageName;

            // TODO(crhodes):
            //	Figure out what to do with roleSource
            //  Figure out how to get PageName shape added.
        }

        public static void CreateToolPage(Visio.Application app, string doc, string page, string shape, string shapeu, String[] args)
        {
            string arg0 = null;
            string backgroundPageName = null;

            if (args.Count() != 2)
            {
                VisioHelper.DisplayInWatchWindow(string.Format("Incorrect Argument Count, expected 2.  Check ShapeSheet"));
            }
            else
            {
                arg0 = args[0];
                backgroundPageName = args[1];
            }

            VisioHelper.DisplayInWatchWindow(string.Format("{0}() arg0:{1}  Background:{2}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name,
                arg0, backgroundPageName));

            // Current shape contains text for new page name.

            Visio.Shape activeShape = app.ActivePage.Shapes[shape];
            VisioHelper.DisplayInWatchWindow(string.Format("  Shape(Name:{0}  Text:{1}", activeShape.Name, activeShape.Text));

            string newPageName = string.Format("{0}{1}{2}", "", "", activeShape.Text);

            Visio.Page newPage = CreatePage(newPageName, backgroundPageName);

            Visio.Hyperlink currentHyperLink = activeShape.AddHyperlink();
            currentHyperLink.SubAddress = newPageName;

            activeShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionHyperlink, 0, 2].Formula = newPageName;

            // TODO(crhodes):
            //  Figure out how to get PageName shape added.
        }

        public static void DisplayLayer(Visio.Page page, string layerName, bool show)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0}(layer:{1} show:{2})",
                System.Reflection.MethodInfo.GetCurrentMethod().Name, layerName, show.ToString()));

            VisioHelper.DisplayInWatchWindow(page.NameU);

            foreach (Visio.Layer layer in page.Layers)
            {
                VisioHelper.DisplayInWatchWindow(string.Format("  {0} - Visible:{1} Print:{2}",
                    layer.Name,
                    layer.CellsC[(short)Visio.VisCellIndices.visLayerVisible].FormulaU.ToString(),
                    layer.CellsC[(short)Visio.VisCellIndices.visLayerPrint].FormulaU.ToString()));

                if (layer.Name == layerName)
                {
                    layer.CellsC[(short)Visio.VisCellIndices.visLayerVisible].FormulaU = (show == true ? "1" : "0");
                }
            }
        }

        public static void GatherInfo(Visio.Page page)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Application app = Globals.ThisAddIn.Application;

            StringBuilder sb = new StringBuilder();

            if (page == null)
            {
                System.Windows.Forms.MessageBox.Show("No ActivePage");
                return;
            }

            //sb.AppendFormat("{0} - {1}\n", "ActivePage.Name", page.Name);
            //sb.AppendFormat("{0} - {1}\n", "ActivePage.NameU", page.NameU);

            //try
            //{
            //    sb.AppendFormat("{0} - {1}\n", "ActivePage.OriginalPage.Name", page.OriginalPage.Name);
            //}
            //catch (Exception ex)
            //{
            //    sb.AppendFormat("{0} - {1}\n", "ActivePage.OriginalPage.Name", "<none>");
            //}

            //sb.AppendFormat("{0} - {1}\n", "ActivePage.AutoSize", page.AutoSize);
            //sb.AppendFormat("{0} - {1}\n", "ActivePage.Background", page.Background);
            //sb.AppendFormat("{0} - {1}\n", "ActivePage.Comments", page.Comments.ToString());
            //sb.AppendFormat("{0} - {1}\n", "ActivePage.Connects", page.Connects.Count);
            //sb.AppendFormat("{0} - {1}\n", "ActivePage.ID", page.ID);
            //sb.AppendFormat("{0} - {1}\n", "ActivePage.Index", page.Index);
            //sb.AppendFormat("{0} - {1}\n", "ActivePage.Layers", page.Layers.Count);
            //sb.AppendFormat("{0} - {1}\n", "ActivePage.LayoutRoutePassive", page.LayoutRoutePassive);

            //sb.AppendFormat("{0} - {1}\n", "ActivePage.PageSheet.Name", page.PageSheet.Name);
            //sb.AppendFormat("{0} - {1}\n", "ActivePage.PrintTileCount", page.PrintTileCount);

            //try
            //{
            //    sb.AppendFormat("{0} - {1}\n", "ActivePage.ReviewerID", page.ReviewerID);
            //}
            //catch (Exception ex)
            //{
            //    sb.AppendFormat("{0} - {1}\n", "ActivePage.ReviewerID", "<none>");
            //}

            //sb.AppendFormat("{0} - {1}\n", "ActivePage.ShapeComments", page.ShapeComments.ToString());
            //sb.AppendFormat("{0} - {1}\n", "ActivePage.Shapes", page.Shapes.Count);
            //sb.AppendFormat("{0} - {1}\n", "ActivePage.Stat", page.Stat);
            //sb.AppendFormat("{0} - {1}\n", "ActivePage.Type", page.Type.ToString());

            //VisioHelper.DisplayInWatchWindow(sb.ToString());


            VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.Name", page.Name));
            VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.NameU", page.NameU));

            try
            {
                VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.OriginalPage.Name", page.OriginalPage.Name));
            }
            catch (Exception ex)
            {
                VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.OriginalPage.Name", "<none>"));
            }

            VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.AutoSize", page.AutoSize));
            VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.Background", page.Background));
            VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.Comments", page.Comments.ToString()));
            VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.Connects", page.Connects.Count));
            VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.ID", page.ID));
            VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.Index", page.Index));
            VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.Layers", page.Layers.Count));
            VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.LayoutRoutePassive", page.LayoutRoutePassive));

            VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.PageSheet.Name", page.PageSheet.Name));
            VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.PrintTileCount", page.PrintTileCount));

            try
            {
                VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.ReviewerID", page.ReviewerID));
            }
            catch (Exception ex)
            {
                VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.ReviewerID", "<none>"));
            }

            VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.ShapeComments", page.ShapeComments.ToString()));
            VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.Shapes", page.Shapes.Count));
            VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.Stat", page.Stat));
            VisioHelper.DisplayInWatchWindow(string.Format("{0} - {1}\n", "ActivePage.Type", page.Type.ToString()));

            //VisioHelper.DisplayInWatchWindow(sb.ToString());
            foreach (Visio.Shape shape in page.Shapes)
            {
                Actions.Visio_Shape.DisplayInfo(shape);
            }
            //System.Windows.Forms.MessageBox.Show(sb.ToString());
        }

        public static void PageChanged(Visio.Page page)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0} Name:>{1}< NameU:>{2}<",
                System.Reflection.MethodInfo.GetCurrentMethod().Name, page.Name, page.NameU));

            SyncPageNames(page);
            UpdatePageNameShapes(page);
        }

        public static void PrintPage()
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                Visio.Application app = Globals.ThisAddIn.Application;
                Visio.Document doc = app.ActiveDocument;
                Visio.Page page = app.ActivePage;

                page.Print();
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void RemoveLayers()
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Application app = Globals.ThisAddIn.Application;
            Visio.Document doc = app.ActiveDocument;
            Visio.Page page = app.ActivePage;

            DeleteLayers(page);
        }

        public static void DeleteLayers(Visio.Page page)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0} Name:>{1}< NameU:>{2}<",
                System.Reflection.MethodInfo.GetCurrentMethod().Name, page.Name, page.NameU));

            try
            {
                // TODO(crhodes):
                // Handle if "Default Layers" page doesn't exist

                Visio.Page layersPage = Globals.ThisAddIn.Application.ActiveDocument.Pages["Default Layers"];
                List<Visio.Shape> layerNames = GetLayerNameShapes(layersPage, LayerNameType.RemovalName);

                foreach (Visio.Shape shape in layerNames)
                {
                    DeleteLayer(page, shape.Text, 0);
                    //foreach (Visio.Layer layer in page.Layers)
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
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void DeleteLayer(Visio.Page page, string layerName, short deleteShapes)
        {
            // TODO(crhodes)
            // may want to pass in a forceUnlock flag defaulted to 0

            try
            {
                Visio.Layer layer = null;

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
                    layer.CellsC[(short)Visio.VisCellIndices.visLayerLock].FormulaU = "0";
                    layer.Delete(deleteShapes);
                }
                else
                {
                    VisioHelper.DisplayInWatchWindow(string.Format("Layer >{0}< does not exist", layerName));
                }
            }
            catch (Exception ex)
            {
                // TODO(crhodes):
                // Decide if what to show this to user.  Layer maybe locked.
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void MovePage(Visio.Page page, string targetDocument)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));
            Visio.Application app = page.Application;
            Visio.Document doc = page.Document;
            string currentDocument = doc.Name;

            try
            {
                Int32 diagramServices = doc.DiagramServicesEnabled;
                doc.DiagramServicesEnabled = (Int32)VisDiagramServices.visServiceAll;

                app.ActiveWindow.Page = page;

                // TODO(crhodes)
                // Need to get all the shapes on the page and ignore the Navigation Shapes

                app.ActiveWindow.SelectAll();
                app.ActiveWindow.Selection.Cut();
                app.Windows.ItemEx[targetDocument].Activate();

                Visio.Page newPage = app.ActiveDocument.Pages.Add();
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
                Log.Error(ex, Common.LOG_CATEGORY);
            }

            VisioHelper.DisplayInWatchWindow("Exit");
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

        //    Dim UndoScopeID1 As Long
        //    UndoScopeID1 = Application.BeginUndoScope("Insert Page")
        //    Dim vsoPage1 As Visio.Page
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

        //    Application.ActiveWindow.Page.Paste

        //    Application.ActivePage.Name = "QMS RDL <DataSources>"

        //    'Restore diagram services
        //    ActiveDocument.DiagramServicesEnabled = DiagramServices

        //End Sub

        public static void SavePage(Visio.Page page)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0} Name:>{1}< NameU:>{2}<",
                System.Reflection.MethodInfo.GetCurrentMethod().Name, page.Name, page.NameU));

            Visio.Application app = Globals.ThisAddIn.Application;

            app.Settings.SetRasterExportResolution(Visio.VisRasterExportResolution.visRasterUseCustomResolution, 150, 150, Visio.VisRasterExportResolutionUnits.visRasterPixelsPerCm);
            //app.Settings.SetRasterExportResolution(Visio.VisRasterExportResolution.visRasterUseCustomResolution, 300, 300, Visio.VisRasterExportResolutionUnits.visRasterPixelsPerCm);
            app.Settings.SetRasterExportSize(Visio.VisRasterExportSize.visRasterFitToSourceSize);
            //app.Settings.SetRasterExportSize(Visio.VisRasterExportSize.visRasterFitToScreenSize);
            //app.Settings.SetRasterExportSize(Visio.VisRasterExportSize.visRasterFitToPrinterSize);
            //app.Settings.SetRasterExportSize(Visio.VisRasterExportSize.visRasterFitToCustomSize, 11.0, 8.5, Visio.VisRasterExportSizeUnits.visRasterInch);

            //app.Settings.SetRasterExportSize(Visio.VisRasterExportSize.visRasterFitToSourceSize, 9.56, 7.47, Visio.VisRasterExportSizeUnits.visRasterInch);
            app.Settings.RasterExportDataFormat = Visio.VisRasterExportDataFormat.visRasterInterlace;
            app.Settings.RasterExportColorFormat = Visio.VisRasterExportColorFormat.visRaster24Bit;
            app.Settings.RasterExportRotation = Visio.VisRasterExportRotation.visRasterNoRotation;
            app.Settings.RasterExportFlip = Visio.VisRasterExportFlip.visRasterNoFlip;
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
                Log.Error(ex, Common.LOG_CATEGORY);
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

        public static void SyncPageNames()
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Application app = Globals.ThisAddIn.Application;
            Visio.Document doc = app.ActiveDocument;
            Visio.Page page = app.ActivePage;

            SyncPageNames(page);
        }

        public static void SyncPageNames(Visio.Page page)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0} Name:>{1}< NameU:>{2}<",
                System.Reflection.MethodInfo.GetCurrentMethod().Name, page.Name, page.NameU));

            try
            {
                Globals.ThisAddIn.Application.EventsEnabled = 0;
                page.NameU = page.Name;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);

            }
            finally
            {
                Globals.ThisAddIn.Application.EventsEnabled = 1;
            }
        }

        public static void ToggleLayerLock(Visio.Application app, string doc, string page, string shape, string shapeu)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0})",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Shape activeShape = app.ActivePage.Shapes[shape];
            string layerName = activeShape.CellsU["Prop.Layer"].ResultStrU[0];

            foreach (Visio.Layer layer in app.ActivePage.Layers)
            {

                VisioHelper.DisplayInWatchWindow(layer.Name);

                if (layer.Name.ToLower() == layerName.ToLower())
                {
                    var currentState = layer.CellsC[(short)Visio.VisCellIndices.visLayerLock].ResultIU;
                    string newState = null;

                    newState = (currentState == 0) ? "1" : "0";

                    //bool state = !bool.Parse(.ToString());
                    layer.CellsC[(short)Visio.VisCellIndices.visLayerLock].Formula = newState;
                    activeShape.CellsU["Prop.Lock"].FormulaU = newState;
                }
            }
        }

        public static void ToggleLayerPrint(Visio.Application app, string doc, string page, string shape, string shapeu)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0})",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Shape activeShape = app.ActivePage.Shapes[shape];
            string layerName = activeShape.CellsU["Prop.Layer"].ResultStrU[0];

            foreach (Visio.Layer layer in app.ActivePage.Layers)
            {

                VisioHelper.DisplayInWatchWindow(layer.Name);

                if (layer.Name.ToLower() == layerName.ToLower())
                {
                    var currentState = layer.CellsC[(short)Visio.VisCellIndices.visLayerPrint].ResultIU;
                    string newState = null;

                    newState = (currentState == 0) ? "1" : "0";

                    //bool state = !bool.Parse(.ToString());
                    layer.CellsC[(short)Visio.VisCellIndices.visLayerPrint].Formula = newState;
                    activeShape.CellsU["Prop.Print"].FormulaU = newState;
                }
            }
        }

        public static void ToggleLayerVisibility(Visio.Application app, string doc, string page, string shape, string shapeu)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0})",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Shape activeShape = app.ActivePage.Shapes[shape];
            string layerName = activeShape.CellsU["Prop.Layer"].ResultStrU[0];

            foreach (Visio.Layer layer in app.ActivePage.Layers)
            {

                VisioHelper.DisplayInWatchWindow(layer.Name);

                if (layer.Name.ToLower() == layerName.ToLower())
                {
                    var currentState = layer.CellsC[(short)Visio.VisCellIndices.visLayerVisible].ResultIU;
                    string newState = null;

                    newState = (currentState == 0) ? "1" : "0";

                    //bool state = !bool.Parse(.ToString());
                    layer.CellsC[(short)Visio.VisCellIndices.visLayerVisible].Formula = newState;
                    activeShape.CellsU["Prop.Visible"].FormulaU = newState;
                }
            }
        }

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

        public static void UpdateGroupNameShapes(Visio.Application app, string doc, string page, string shape, string shapeu, String[] args)
        {
            string mode = null;

            if (args.Count() != 1)
            {
                //VisioHelper.DisplayInWatchWindow(string.Format("Incorrect Argument Count, expected 3.  Check ShapeSheet"));
            }
            else
            {
                mode = args[0];
            }

            VisioHelper.DisplayInWatchWindow(string.Format("{0}({1})  mode:{2}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name, page,
                mode));

            Visio.Page currentPage = app.ActivePage;
            Visio.Shape colorSelectorShape = currentPage.Shapes[shape];
            string colorSelectorGroupName = null;

            if (colorSelectorShape.CellExists["Prop.GroupName", 0] != 0)
            {
                colorSelectorGroupName = colorSelectorShape.Cells["Prop.GroupName"].ResultStrU[0];
            }
            else
            {
                VisioHelper.DisplayInWatchWindow(string.Format("Cannot locate Prop.GroupName.  Check ShapeSheet"));
                return;
            }

            Visio.Cell userCell;
            Visio.Cell propCell;
            Visio.Cell propCell2;
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

            VisioHelper.DisplayInWatchWindow(string.Format("userColor:{0}-{1} propColor:{2}-{3} {4}", userColor, userColorS, propColor, propColorS, propColor2S));

            // Now walk the shapes on the page looking for shapes with a matching GroupName

            string groupName = string.Empty;
            short isSelectorTool = 0;
            short hasGroupName = 0;

            foreach (Visio.Shape pageShape in currentPage.Shapes)
            {
                VisioHelper.DisplayInWatchWindow(string.Format("shape NameID:({0})", pageShape.NameID));
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

                            VisioHelper.DisplayInWatchWindow(string.Format("   groupName:({0})  isSelectorTool:({1})", groupName, isSelectorTool));

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
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }

        public static void UpdateHasColorTagsShapes(Microsoft.Office.Interop.Visio.Application app, string doc, string page, string shape, string shapeu, String[] args)
        {
            int levels = 0;

            if (args.Count() != 1)
            {
                //VisioHelper.DisplayInWatchWindow(string.Format("Incorrect Argument Count, expected 3.  Check ShapeSheet"));
            }
            else
            {
                levels = int.Parse(args[0]);
            }

            VisioHelper.DisplayInWatchWindow(string.Format("{0}({1})  levels:{2}",
                System.Reflection.MethodInfo.GetCurrentMethod().Name, page,
                levels));

            Visio.Page currentPage = app.ActivePage;
            Visio.Shape colorTagShape = currentPage.Shapes[shape];

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
                VisioHelper.DisplayInWatchWindow(string.Format("Cannot locate Prop.TagName.  Check ShapeSheet"));
                return;
            }

            if (colorTagShape.CellExists["Prop.ForegroundColor", 0] != 0)
            {
                foregroundColor = colorTagShape.Cells["Prop.ForegroundColor"].ResultStrU[0];
            }
            else
            {
                VisioHelper.DisplayInWatchWindow(string.Format("Cannot locate Prop.ForegroundColor.  Check ShapeSheet"));
                return;
            }

            if (colorTagShape.CellExists["Prop.BackgroundColor", 0] != 0)
            {
                backgroundColor = colorTagShape.Cells["Prop.BackgroundColor"].ResultStrU[0];
            }
            else
            {
                VisioHelper.DisplayInWatchWindow(string.Format("Cannot locate Prop.BackgroundColor.  Check ShapeSheet"));
                return;
            }

            if (colorTagShape.CellExists["Prop.Pattern", 0] != 0)
            {
                pattern = colorTagShape.Cells["Prop.Pattern"].ResultStrU[0];
            }
            else
            {
                VisioHelper.DisplayInWatchWindow(string.Format("Cannot locate Prop.Pattern.  Check ShapeSheet"));
                return;
            }

            if (colorTagShape.CellExists["Prop.IsVisible", 0] != 0)
            {
                isVisible = colorTagShape.Cells["Prop.IsVisible"].ResultStrU[0];
            }
            else
            {
                VisioHelper.DisplayInWatchWindow(string.Format("Cannot locate Prop.IsVisible.  Check ShapeSheet"));
                return;
            }

            foreach (Visio.Shape pageShape in currentPage.Shapes)
            {
                VisioHelper.DisplayInWatchWindow(string.Format("shape NameID:({0})", pageShape.NameID));

                try
                {
                    var hasColorTags = pageShape.CellExistsU["User.HasColorTags", 0];    // 0 is Local and Inherited, 1 is Local only 

                    VisioHelper.DisplayInWatchWindow(string.Format("shape {0}  hasColorTags:{1})",
                        pageShape.Name, hasColorTags));

                    if (hasColorTags != 0)
                    {
                        foreach (Visio.Shape subShape in pageShape.Shapes)
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
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }

        public static void UpdateLayer(Microsoft.Office.Interop.Visio.Application app, string doc, string page, string shape, string shapeu)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0})",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            Visio.Shape activeShape = app.ActivePage.Shapes[shape];
            string layerName = activeShape.CellsU["Prop.Layer"].ResultStrU[0];          

            foreach (Visio.Layer layer in app.ActivePage.Layers)
            {
                try
                {
                    VisioHelper.DisplayInWatchWindow(layer.Name);

                    if (layer.Name.ToLower() == layerName.ToLower())
                    {
                        if (activeShape.CellExistsU["Prop.Visible", 0] != 0)
                        {
                            layer.CellsC[(short)Visio.VisCellIndices.visLayerVisible].FormulaU = activeShape.CellsU["Prop.Visible"].ResultStrU[0];
                        }

                        if (activeShape.CellExistsU["Prop.Lock", 0] != 0)
                        {
                            layer.CellsC[(short)Visio.VisCellIndices.visLayerLock].FormulaU = activeShape.CellsU["Prop.Lock"].ResultStrU[0];
                        }

                        if (activeShape.CellExistsU["Prop.Print", 0] != 0)
                        {
                            layer.CellsC[(short)Visio.VisCellIndices.visLayerPrint].FormulaU = activeShape.CellsU["Prop.Print"].ResultStrU[0];
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }

        public static void UpdatePageNameShapes(Visio.Page page)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0} Name:>{1}< NameU:>{2}<",
                System.Reflection.MethodInfo.GetCurrentMethod().Name, page.Name, page.NameU));

            foreach (Visio.Shape shape in page.Shapes)
            {
                Actions.Visio_Shape.UpdatePageNameShape(shape, page.Name);                
            }
        }

        #endregion

        #region Private Methods

        private static List<Visio.Shape> GetLayerNameShapes(Visio.Page page, LayerNameType nameType = LayerNameType.AllNames)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0}({1})",
                System.Reflection.MethodInfo.GetCurrentMethod().Name, page.NameU));

            List<Visio.Shape> layerNames = new List<Visio.Shape>();

            foreach (Visio.Shape shape in page.Shapes)
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
                            VisioHelper.DisplayInWatchWindow(string.Format("Unknown LayerNameype:{0}",
                                nameType));
                            break;
                    }
                }
            }

            return layerNames;
        }

        private static List<Visio.Shape> GetNavigationLinks(Visio.Page page)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0}({1})",
                System.Reflection.MethodInfo.GetCurrentMethod().Name, page.NameU));

            List<Visio.Shape> navigationLinks = new List<Visio.Shape>();

            foreach (Visio.Shape shape in page.Shapes)
            {
                var isNavigationLink = shape.CellExists["User.IsNavigationLink", 0];

                navigationLinks.Add(shape);
            }

            return navigationLinks;
        }

        private static string GetPageSaveName(Visio.Page page)
        {
            string pageName = VNC.AddinHelper.Util.SafeFileName(page.NameU);
            string documentName = VNC.AddinHelper.Util.SafeFileName(page.Application.ActiveDocument.Name);

            // TODO(crhodes):
            // Do more fancy stuff so it is easier to find the file later

            pageName = string.Format(@"C:\temp\VisioExport\{0}-{1}.png", documentName, pageName);

            return pageName;
        }

        private static void RemoveNavigationLinks(Visio.Page page)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0} Name:>{1}< NameU:>{2}<",
                System.Reflection.MethodInfo.GetCurrentMethod().Name, page.Name, page.NameU));

            List<Visio.Shape> navigationLinks = GetNavigationLinks(page);

            try
            {
                foreach (Visio.Shape shape in navigationLinks)
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
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

         #endregion

    }
}
