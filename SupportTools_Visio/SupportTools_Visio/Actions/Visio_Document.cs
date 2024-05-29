using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Microsoft.Office.Interop.Visio;

using VNC;

using Visio = Microsoft.Office.Interop.Visio;
using VisioHelper = VNC.AddinHelper.Visio;

namespace SupportTools_Visio.Actions
{
    class Visio_Document
    {

        const double cTOC_Initial_xLoc = 1;
        const double cTOC_Initial_yLoc = 8.0;
        const double cTOC_Link_Width = 2.5;
        const double cTOC_Link_Height = 0.125;

        const int cTOC_Link_FontSize = 10;

        const int cTOC_MaxItemsInColumn = 25;
        const double cTOC_Offset_Row = 0.25;
        const double cTOC_Offset_Column = 2.5;

        const double cTOC_Page_Initial_xLoc = 9.75;

        const double cTOC_PageLink_Initial_yLoc = 8.125;
        const double cTOC_PageLink_Width = 1.0;
        const double cTOC_PageLink_Height = 0.125;

        const int cTOC_PageLink_FontSize = 6;

        #region Main Methods

        public static void AddDefaultLayers()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Document doc = app.ActiveDocument;

            if (doc == null)
            {
                System.Windows.Forms.MessageBox.Show("No ActiveDocument");
                return;
            }

            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("AddDefaultLayers");

            foreach (Visio.Page page in doc.Pages)
            {
                Visio_Page.AddDefaultLayers(page);
            }

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
        }

        public static void AddHeader()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Document doc = app.ActiveDocument;

            // TODO: Add some stuff to read from config file with a dialog to default the selection

            doc.HeaderLeft = "";
            doc.HeaderCenter = "";
            doc.HeaderRight = "";

            //doc.FooterLeft = "&f&e";
            //doc.FooterCenter = "";
            //doc.FooterRight = "&d &p-&P";

            var font = doc.HeaderFooterFont;

            font.Size = (decimal)8;

            doc.HeaderFooterFont = font;

            var size = doc.HeaderFooterFont.Size;

        }

        public static void AddFooter()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Document doc = app.ActiveDocument;

            // TODO: Add some stuff to read from config file with a dialog to default the selection

            //doc.HeaderLeft = "";
            //doc.HeaderCenter = "&n";
            //doc.HeaderRight = "";

            doc.FooterLeft = "&f&e";
            doc.FooterCenter = "";
            doc.FooterRight = "&d &p-&P";

            var font = doc.HeaderFooterFont;

            font.Size = (decimal)8;

            doc.HeaderFooterFont = font;

            var size = doc.HeaderFooterFont.Size;

            //doc.HeaderMargin = 0.13;
            //doc.FooterMargin = 0.13;
            doc.HeaderMargin[VisUnitCodes.visDrawingUnits] = 0.13;
            doc.FooterMargin[VisUnitCodes.visDrawingUnits] = 0.13;
        }

        public static void AddNavigationLinks()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("AddNavigationLinks");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Document doc = app.ActiveDocument;

            foreach (Visio.Page page in doc.Pages)
            {
                Visio_Page.AddNavigationLinks(page);
            }

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
        }

        public static void AutoSizePagesOff()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("AutoSizePagesOff");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Document doc = app.ActiveDocument;

            foreach (Visio.Page page in doc.Pages)
            {
                Visio_Page.AutoSizePageOff(page);
            }

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
        }

        public static void AutoSizePagesOn()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("AutoSizePagesOn");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Document doc = app.ActiveDocument;

            foreach (Visio.Page page in doc.Pages)
            {
                Visio_Page.AutoSizePageOn(page);
            }

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
        }

        public static void CreateTableOfContents()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Page pageTOC = CreateTOCPage();

            // Use Navigation Links instead.

            //foreach (Visio.Page page in Globals.ThisAddIn.Application.ActiveDocument.Pages)
            //{
            //    if ( ! page.NameU.Equals("Table of Contents"))
            //    {

            //        //AddTOCLinkToPage(page);
            //    }
            //}

            // Should drive this off a calculation based on page size, # of pages, etc..  Hack it for now.

            double xLoc = cTOC_Initial_xLoc;
            double yLoc = cTOC_Initial_yLoc;

            int columnCount = 0;

            foreach (Visio.Page page in Globals.ThisAddIn.Application.ActiveDocument.Pages)
            {
                if ( ! page.NameU.Equals("Table of Contents"))
                {
                    AddPageLinkToTOCPage(pageTOC, page, xLoc, yLoc);
                    yLoc -= cTOC_Offset_Row;
                    columnCount++;

                    if (columnCount > cTOC_MaxItemsInColumn)
                    {
                        xLoc += cTOC_Offset_Column;
                        yLoc = cTOC_Initial_yLoc;
                        columnCount = 0;
                    }
                }
            }
        }

        public static Visio.Page CreateTOCPage()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Page page = null;

            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("GenerateTOCPage");

            try
            {
                page = Globals.ThisAddIn.Application.ActiveDocument.Pages["Table of Contents"];
                // We found a page, delete it.  Not much luck iterating across shapes and clearing page - See ClearPage()

                page.Delete(0);
                //ClearPage(page);
                // Need to delete all the stuff.
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);

            }

            page = Globals.ThisAddIn.Application.ActiveDocument.Pages.Add();

            page.NameU = "Table of Contents";
            page.Background = 0;
            page.Index = 1;

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);

            return page;
        }

        public static void DeletePages()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;
            Visio.Document doc = app.ActiveDocument;
            Visio.Page page = app.ActivePage;

            foreach (Visio.Shape shape in page.Shapes)
            {
                VisioHelper.DisplayInWatchWindow(string.Format("Name: {0}  Text: {1}", shape.Name, shape.Text));
                try
                {
                    short renumberPages = 0;    // Do not renumber default named pages
                    doc.Pages[shape.Text].Delete(renumberPages);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }

        public static void DisplayInfo()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            StringBuilder sb = new StringBuilder();

            Visio.Document doc = app.ActiveDocument;

            if (doc == null)
            {
                System.Windows.Forms.MessageBox.Show("No ActiveDocument");
                return;
            }

            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Name", doc.Name);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Cateogory", doc.Category);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.ClasssID", doc.ClassID);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Comments", doc.Comments);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Company", doc.Company);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.CompatibilityMode", doc.CompatibilityMode);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Creator", doc.Creator);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.DefaultFillStyle", doc.DefaultFillStyle);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.DefaultGuideStyle", doc.DefaultGuideStyle);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.DefaultLineStyle", doc.DefaultLineStyle);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.DefaultSavePath", doc.DefaultSavePath);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.DefaultStyle", doc.DefaultStyle);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.DefaultTextStyle", doc.DefaultTextStyle);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Description", doc.Description);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.DefaultGuideStyle", doc.DefaultGuideStyle);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.DocumentSheet.Name", doc.DocumentSheet.Name);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.DynamicGridEnabled", doc.DynamicGridEnabled);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.EditorCount", doc.EditorCount);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.FooterCenter", doc.FooterCenter);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.FooterLeft", doc.FooterLeft);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.FooterMargin", doc.FooterMargin);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.FooterRight", doc.FooterRight);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.FullName", doc.FullName);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.GestureFormatSheet.Name", doc.GestureFormatSheet.Name);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.GlueEnabled", doc.GlueEnabled);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.HeaderCenter", doc.HeaderCenter);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.HeaderFooterColor", doc.HeaderFooterColor);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.HeaderFooterFont", doc.HeaderFooterFont);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.HeaderLeft", doc.HeaderLeft);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.HeaderMargin", doc.HeaderMargin);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.HeaderRight", doc.HeaderRight);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.HyperlinkBase", doc.HyperlinkBase);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.ID", doc.ID);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Index", doc.Index);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.InPlace", doc.InPlace);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Keywords", doc.Keywords);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.LeftMargin", doc.LeftMargin);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Manager", doc.Manager);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Masters.Count", doc.Masters.Count);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Pages.Count", doc.Pages.Count);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Path", doc.Path);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.PrintCenteredH", doc.PrintCenteredH);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.PrintCenteredV", doc.PrintCenteredV);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Printer", doc.Printer);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.PrintFitOnPages", doc.PrintFitOnPages);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.PrintLandscape", doc.PrintLandscape);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.PrintPagesAcross", doc.PrintPagesAcross);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.PrintPagesDown", doc.PrintPagesDown);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.ProgID", doc.ProgID);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.ReadOnly", doc.ReadOnly);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.RemovePersonalInformation", doc.RemovePersonalInformation);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.RightMargin", doc.RightMargin);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Saved", doc.Saved);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.SnapEnabled", doc.SnapEnabled);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Stat", doc.Stat);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Subject", doc.Subject);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Template", doc.Template);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.TimeCreated", doc.TimeCreated.ToString());
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.TimeEdited", doc.TimeEdited.ToString());
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.TimePrinted", doc.TimePrinted.ToString());
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.TimeSaved", doc.TimeSaved.ToString());
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Title", doc.Title);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.TopMargin", doc.TopMargin);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.UndoEnabled", doc.UndoEnabled);
            sb.AppendFormat("{0} - {1}\n", "ActiveDocument.Version", doc.Version);

            VisioHelper.DisplayInWatchWindow(sb.ToString());        }

        public static void DisplayLayer(string layerName, bool show)
        {
            VisioHelper.DisplayInWatchWindow(string.Format("{0}(layer:{1} show:{2})",
                MethodBase.GetCurrentMethod().Name, layerName, show.ToString()));

            foreach (Visio.Page page in Globals.ThisAddIn.Application.ActiveDocument.Pages)
            {
                Visio_Page.DisplayLayer(page, layerName, show);
            }
        }

        public static void DisplayPageNames()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Document doc = app.ActiveDocument;

            try
            {
                foreach (Visio.Page page in doc.Pages)
                {
                    VisioHelper.DisplayInWatchWindow(string.Format("Page {0} Name:>{1:30}< NameU:>{2:30}<", 
                        page.Name.Equals(page.NameU) == true ? " " : "?",
                        page.Name, page.NameU));
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static List<Visio.Shape> GetNavigationLinks()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            List<Visio.Shape> navLinks = new List<Visio.Shape>();

            Visio.Page linkPage = Globals.ThisAddIn.Application.ActiveDocument.Pages["Navigation Links"];

            foreach (Visio.Shape shape in linkPage.Shapes)
            {
                navLinks.Add(shape);
            }

            return navLinks;
        }

        public static void MovePages(string targetDocument)
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;
            Visio.Document doc = app.ActiveDocument;
            Visio.Page page = app.ActivePage;

            foreach (Visio.Shape shape in page.Shapes)
            {
                VisioHelper.DisplayInWatchWindow(string.Format("Name: {0}  Text: {1}", shape.Name, shape.Text));
                try
                {
                    Visio_Page.MovePage(doc.Pages[shape.Text], targetDocument);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }

        public static void PrintPages()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            try
            {
                Visio.Application app = Globals.ThisAddIn.Application;
                Visio.Document doc = app.ActiveDocument;
                Visio.Page page = app.ActivePage;

                foreach (Visio.Shape shape in page.Shapes)
                {
                    VisioHelper.DisplayInWatchWindow(string.Format("Name: {0}  Text: >{1}<", shape.Name, shape.Text));

                    //var bar = shape.Hyperlink;
                    //var hyperLinks = shape.Hyperlinks;

                    //foreach (Visio.Hyperlink hyperlink in hyperLinks)
                    //{
                    //    var foo = hyperlink.Address;
                    //}

                    if (shape.Hyperlinks.Count > 0)
                    {
                        if (shape.Hyperlink.SubAddress.Length > 0)
                        {
                            VisioHelper.DisplayInWatchWindow(string.Format("Hyperlink: >{0}<", shape.Hyperlink.SubAddress));
                            doc.Pages[shape.Hyperlink.SubAddress].Print();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void RemoveLayers()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;
            Visio.Document doc = app.ActiveDocument;

            foreach (Visio.Page page in doc.Pages)
            {
                Visio_Page.DeleteLayers(page);
            }
        }

        public static void RenamePages(string searchExpression, string replacementExpression,
            RegexOptions regexOptions = RegexOptions.None)
        {
            VNC.Log.Trace("", Common.LOG_CATEGORY, 0);

            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;
            Visio.Document doc = app.ActiveDocument;
            Visio.Page page = app.ActivePage;

            Regex regex = new Regex(searchExpression, regexOptions);

            foreach (Visio.Shape shape in page.Shapes)
            {
                VisioHelper.DisplayInWatchWindow(string.Format("Name: {0}  Text: {1}", shape.Name, shape.Text));
                try
                {
                    string newPageName = regex.Replace(doc.Pages[shape.Text].Name, replacementExpression);

                    doc.Pages[shape.Text].Name = newPageName;
                }
                catch (Exception ex)
                {
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }

        public static void SavePages()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;
            Visio.Document doc = app.ActiveDocument;
            Visio.Page page = app.ActivePage;

            foreach (Visio.Shape shape in page.Shapes)
            {
                VisioHelper.DisplayInWatchWindow(string.Format("Name: {0}  Text: {1}", shape.Name, shape.Text));
                try
                {
                    Visio_Page.SavePage(doc.Pages[shape.Text]);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }

        public static void SortAllPages()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Document doc = app.ActiveDocument;

            System.Collections.SortedList sortedPages = new System.Collections.SortedList();
            //SortedList<string, string> sortedPages = new SortedList<string, string>();
            int index = 0;
            bool hasTOCPage = false;

            VisioHelper.DisplayInWatchWindow(string.Format("Document({0})", doc.Name));

            try
            {
                foreach (Visio.Page page in doc.Pages)
                {
                    VisioHelper.DisplayInWatchWindow(string.Format("Page({0} IsBackground{1})", page.NameU, page.Background));

                    if ( ! page.NameU.Equals("Table of Contents"))
                    {
                        if (page.Background == 0)
                        {
                            sortedPages.Add(page.NameU, page.NameU);
                            index++;
                        }
                    }
                    else
                    {
                        hasTOCPage = true;
                    }

                    //sortedPages.Add(index++, page.NameU);
                }

                // If we found a TOC page, start pages off at postion 2, else, postion 1

                int offset = hasTOCPage ? 2 : 1;

                for (int i = 0; i < index; i++)
                {
                    VisioHelper.DisplayInWatchWindow(string.Format("Moving Page({0})", sortedPages.GetByIndex(i)));
                    doc.Pages.ItemU[sortedPages.GetKey(i)].Index = (short)(i + offset);
                    //Application.ActiveDocument.Pages.ItemU("Page-2").Index = 3
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void SyncPageNames()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Document doc = app.ActiveDocument;

            List<Visio.Page> pagesToUpdate = new List<Microsoft.Office.Interop.Visio.Page>();

            VisioHelper.DisplayInWatchWindow(string.Format("Document({0})", doc.Name));

            foreach (Visio.Page page in doc.Pages)
            {
                if (! page.Name.Equals(page.NameU))
                {
                    pagesToUpdate.Add(page);
                }
            }

            foreach (Visio.Page page in pagesToUpdate)
            {
                Visio_Page.SyncPageNames(page);
            }
        }

        public static void UpdatePageNameShapes()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Document doc = app.ActiveDocument;

            foreach (Visio.Page page in doc.Pages)
            {
                Visio_Page.UpdatePageNameShapes(page);
            }
        }

        #endregion

        #region Private Methods

        private static void AddPageLinkToTOCPage(Visio.Page pageTOC, Visio.Page page, double xLoc, double yLoc)
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("AddPageLinkToTOCPage");

            Visio.Shape pageLinkShape = pageTOC.DrawRectangle(xLoc, yLoc, xLoc + cTOC_Link_Width, yLoc + cTOC_Link_Height);

            pageLinkShape.TextStyle = "Normal";
            pageLinkShape.LineStyle = "Text Only";
            pageLinkShape.FillStyle = "Text Only";
            pageLinkShape.Characters.Begin = 0;
            pageLinkShape.Characters.End = 0;
            pageLinkShape.Text = page.NameU;
            pageLinkShape.Characters.set_CharProps((short)Visio.VisCellIndices.visCharacterSize, cTOC_Link_FontSize);

            Visio_Shape.AddHyperLink(pageLinkShape, page.NameU);

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
        }

        private static void AddTOCLinkToPage(Visio.Page page)
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            if (page.Background != 0)
            {
                // Skip background pages
            	return;
            }

            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("AddTOCLinkToPage");

            // Clear out any existing link.

            foreach (Visio.Shape shape in page.Shapes)
            {
                if (shape.Text == "Table of Contents" || shape.Name == "TOCLink")
                {
                    shape.Delete();
                }
            }

            Visio.Shape tocShape = page.DrawRectangle(
                cTOC_Page_Initial_xLoc, cTOC_PageLink_Initial_yLoc,
                cTOC_Page_Initial_xLoc + cTOC_PageLink_Width, cTOC_PageLink_Initial_yLoc + cTOC_PageLink_Height);

            tocShape.Name = "TOCLink";

            tocShape.Text = "Table of Contents";
            tocShape.TextStyle = "Normal";
            tocShape.LineStyle = "Text Only";
            tocShape.FillStyle = "Text Only";
            tocShape.Characters.Begin = 0;
            tocShape.Characters.End = 0;
            tocShape.Characters.set_CharProps((short)Visio.VisCellIndices.visCharacterSize, 6);

            Visio.Hyperlink hlink = tocShape.Hyperlinks.Add();
            // hlink.Name = "do we need a name?";
            hlink.SubAddress = "Table of Contents";

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
        }

        private static void ClearPage(Visio.Page page)
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            try
            {
                for (int i = page.Shapes.Count - 1; i >= 0; i--)
                {
                    page.Shapes[i].Delete();
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Shapes on Page: {0}", page.Shapes.Count));

        }
        public static void AddArchitectureBasePages()
        {
            VisioHelper.DisplayInWatchWindow($"{MethodBase.GetCurrentMethod().Name}()");

            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Page page = null;
            Visio.Document archStencil = null;
            Visio.Master archMaster = null;

            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("AddArchitectureBasePages");

            var architecturePages = new string[] { 
                "Clean Architecture 1", 
                "Clean Architecture 2", 
                "Clean Architecture 3", 
                "Clean Architecture 4" };

            // Delete any existing Architecture Base Pages

            foreach (var  existingPage in architecturePages)
            {
                try
                {
                    page = app.ActiveDocument.Pages[existingPage];
                    // We found a page, delete it.  Not much luck iterating across shapes and clearing page - See ClearPage()

                    page.Delete(0);
                    //ClearPage(page);
                    // Need to delete all the stuff.
                }
                catch (Exception ex)
                {
                    // Maybe log that we found an existing page
                    //Log.Error(ex, Common.LOG_CATEGORY);
                }
            }

            var archStencilName = "API.vssx";

            try
            {
                archStencil = app.Documents[archStencilName];
            }
            catch (Exception ex)
            {
                // Open Stencil
                app.Documents.OpenEx(archStencilName, (short)Visio.VisOpenSaveArgs.visOpenRO + (short)Visio.VisOpenSaveArgs.visOpenDocked);

                archStencil = app.Documents[archStencilName];
            }

            // Assume 11 x 8.5 Landscape page

            try
            {
                page = Globals.ThisAddIn.Application.ActiveDocument.Pages.Add();
                archMaster = archStencil.Masters["Clean Arch 1 - Page Base"];

                page.NameU = "Clean Architecture 1";
                page.Background = 1;
                page.Drop(archMaster, 5.5, 4.325);

                page = Globals.ThisAddIn.Application.ActiveDocument.Pages.Add();
                archMaster = archStencil.Masters["Clean Arch 2 - Page Base"];

                page.NameU = "Clean Architecture 2";
                page.Background = 1;
                page.Drop(archMaster, 5.5, 4.325);

                page = Globals.ThisAddIn.Application.ActiveDocument.Pages.Add();
                archMaster = archStencil.Masters["Clean Arch 3 - Page Base"];

                page.NameU = "Clean Architecture 3";
                page.Background = 1;
                page.Drop(archMaster, 5.5, 4.325);

                page = Globals.ThisAddIn.Application.ActiveDocument.Pages.Add();
                archMaster = archStencil.Masters["Clean Arch 4 - Page Base"];

                page.NameU = "Clean Architecture 4";
                page.Background = 1;
                page.Drop(archMaster, 5.5, 4.325);
            }
            catch (Exception ex)
            {
                
            }

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
        }
        public static void CreatePluralSightCourseFileFromShape(Visio.Application app, string doc, string page, string shape, string shapeu, string[] array)
        {
            int i = 0;

            Visio.Page currentPage = app.ActivePage;
            Visio.Document currentDocument = app.ActiveDocument;
            Visio.Shape activeShape = app.ActivePage.Shapes[shape];
            string shapePageName = "COURSENAME";
            string shapeAuthor = "AUTHOR";

            try
            {
                if (activeShape.CellExistsU["Prop.PageName", 0] != 0)
                {
                    shapePageName = activeShape.CellsU["Prop.PageName"].ResultStrU[Visio.VisUnitCodes.visUnitsString];
                }

                if (activeShape.CellExistsU["Prop.Author", 0] != 0)
                {
                    shapeAuthor = activeShape.CellsU["Prop.Author"].ResultStrU[Visio.VisUnitCodes.visUnitsString];
                }

                string templateName = "CHR Notes - PluralSight Course - Subject - Author.vstx";
                app.Documents.AddEx(templateName, Visio.VisMeasurementSystem.visMSDefault, 0, 0);

                app.ActivePage.Drop(activeShape, (short)5.5, (short)4.25);

                string fileName = $"CHR Notes - PluralSight Course - {shapePageName} - {shapeAuthor}.vsdx";

                SaveFileDialog saveFileDiaglog = new SaveFileDialog();

                saveFileDiaglog.FileName = fileName;
                saveFileDiaglog.InitialDirectory = @"B:\CHR Notes\PluralSight";

                DialogResult result = saveFileDiaglog.ShowDialog();

                fileName = saveFileDiaglog.FileName;

                app.ActiveDocument.SaveAsEx(fileName, (short)Visio.VisOpenSaveArgs.visSaveAsWS);
            }
            catch (Exception ex)
            {
                
            }
        }


        #endregion
    }
}
