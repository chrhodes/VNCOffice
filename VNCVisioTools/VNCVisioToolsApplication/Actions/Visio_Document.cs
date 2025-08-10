using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

using System.Windows.Forms;

using VNC;

using MSVisio = Microsoft.Office.Interop.Visio;

namespace VNCVisioToolsApplication.Actions
{
    public class Visio_Document
    {
        #region Enums, Fields, Properties, Structures

        const double cTOC_Initial_xLoc = 1;
        const double cTOC_Initial_yLoc = 8.0;
        const int cTOC_Link_FontSize = 10;
        const double cTOC_Link_Height = 0.125;
        const double cTOC_Link_Width = 2.5;
        const int cTOC_MaxItemsInColumn = 25;
        const double cTOC_Offset_Column = 2.5;
        const double cTOC_Offset_Row = 0.25;
        const double cTOC_Page_Initial_xLoc = 9.75;

        const int cTOC_PageLink_FontSize = 6;
        const double cTOC_PageLink_Height = 0.125;
        const double cTOC_PageLink_Initial_yLoc = 8.125;
        const double cTOC_PageLink_Width = 1.0;
        #endregion

        #region Events

        #endregion

        #region Action Handlers

        #endregion

        #region Helpers

        #endregion

        #region Main Methods

        public static void AddDefaultLayers()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Document doc = app.ActiveDocument;

            if (doc == null)
            {
                MessageBox.Show("No ActiveDocument");
                return;
            }

            int undoScope = Common.VisioApplication.BeginUndoScope("AddDefaultLayers");

            foreach (MSVisio.Page page in doc.Pages)
            {
                Visio_Page.AddDefaultLayers(page);
            }

            Common.VisioApplication.EndUndoScope(undoScope, true);
        }



        public static void AddFooter()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Document doc = app.ActiveDocument;

            if (doc == null)
            {
                MessageBox.Show("No ActiveDocument");
                return;
            }

            // TODO: Add some stuff to read from config file with a dialog to default the selection

            int undoScope = Common.VisioApplication.BeginUndoScope("AddFooter");

            doc.FooterLeft = "&f&e";
            doc.FooterCenter = "";
            doc.FooterRight = "&d &p-&P";

            //var a = doc.Fonts;

            //stdole.StdFontClass newFont = new stdole.StdFontClass();

            //newFont.Na


            //stdole.IFontDisp hff = doc.HeaderFooterFont;

            //Method not found: 'stdole.IFontDisp Microsoft.Office.Interop.Visio.IVDocument.get_HeaderFooterFont()'.

            var font = doc.HeaderFooterFont;

            font.Name = "Calibri";
            font.Bold = false;
            font.Italic = false;
            font.Underline = false;
            font.Strikethrough = false;

            font.Size = (decimal)6;

            doc.HeaderFooterFont = font;

            //var size = doc.HeaderFooterFont.Size;

            // HACK(crhodes)
            // This is from Recorded Macro

            //Dim oFont2 As StdFont
            //Set oFont2 = Application.ActiveDocument.HeaderFooterFont
            //oFont2.Name = "Calibri"
            //oFont2.Bold = False
            //oFont2.Italic = False
            //oFont2.Underline = False
            //oFont2.Strikethrough = False
            //oFont2.Weight = 400
            //oFont2.Size = 12
            //oFont2.CharSet = 0
            //Set Application.ActiveDocument.HeaderFooterFont = oFont2


            doc.FooterMargin[MSVisio.VisUnitCodes.visInches] = 0.13;

            Common.VisioApplication.EndUndoScope(undoScope, true);
        }

        public static void AddHeader()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Document doc = app.ActiveDocument;

            if (doc == null)
            {
                MessageBox.Show("No ActiveDocument");
                return;
            }

            // TODO: Add some stuff to read from config file with a dialog to default the selection

            int undoScope = Common.VisioApplication.BeginUndoScope("AddHeader");

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

            doc.HeaderMargin[MSVisio.VisUnitCodes.visInches] = 0.13;

            Common.VisioApplication.EndUndoScope(undoScope, true);
        }

        public static void AddNavigationLinks()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            int undoScope = Common.VisioApplication.BeginUndoScope("AddNavigationLinks");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Document doc = app.ActiveDocument;

            foreach (MSVisio.Page page in doc.Pages)
            {
                Visio_Page.AddNavigationLinks(page);
            }

            Common.VisioApplication.EndUndoScope(undoScope, true);
        }

        public static void AutoSizePagesOff()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            int undoScope = Common.VisioApplication.BeginUndoScope("AutoSizePagesOff");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Document doc = app.ActiveDocument;

            foreach (MSVisio.Page page in doc.Pages)
            {
                Visio_Page.AutoSizePageOff(page);
            }

            Common.VisioApplication.EndUndoScope(undoScope, true);
        }

        public static void AutoSizePagesOn()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            int undoScope = Common.VisioApplication.BeginUndoScope("AutoSizePagesOn");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Document doc = app.ActiveDocument;

            foreach (MSVisio.Page page in doc.Pages)
            {
                Visio_Page.AutoSizePageOn(page);
            }

            Common.VisioApplication.EndUndoScope(undoScope, true);
        }

        public static void CreateTableOfContents()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Page pageTOC = CreateTOCPage();

            // Use Navigation Links instead.

            //foreach (MSVisio.Page page in Common.VisioApplication.ActiveDocument.Pages)
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

            foreach (MSVisio.Page page in Common.VisioApplication.ActiveDocument.Pages)
            {
                if (!page.NameU.Equals("Table of Contents"))
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

        public static MSVisio.Page CreateTOCPage()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Page page = null;

            int undoScope = Common.VisioApplication.BeginUndoScope("GenerateTOCPage");

            try
            {
                page = Common.VisioApplication.ActiveDocument.Pages["Table of Contents"];
                // We found a page, delete it.  Not much luck iterating across shapes and clearing page - See ClearPage()

                page.Delete(0);
                //ClearPage(page);
                // Need to delete all the stuff.
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force:true);

            }

            page = Common.VisioApplication.ActiveDocument.Pages.Add();

            page.NameU = "Table of Contents";
            page.Background = 0;
            page.Index = 1;

            Common.VisioApplication.EndUndoScope(undoScope, true);

            return page;
        }

        public static void DeletePages()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;
            MSVisio.Document doc = app.ActiveDocument;
            MSVisio.Page page = app.ActivePage;

            foreach (MSVisio.Shape shape in page.Shapes)
            {
                Common.WriteToDebugWindow(string.Format("Name: {0}  Text: {1}", shape.Name, shape.Text));
                try
                {
                    short renumberPages = 0;    // Do not renumber default named pages
                    doc.Pages[shape.Text].Delete(renumberPages);
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow(ex.ToString(), force:true);
                }
            }
        }

        public static void DisplayInfo()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Document doc = app.ActiveDocument;

            if (doc == null)
            {
                MessageBox.Show("No ActiveDocument");
                return;
            }

            Common.WriteToDebugWindow($"ActiveDocument.Name - {doc.Name}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Category - {doc.Category}\n");
            Common.WriteToDebugWindow($"ActiveDocument.ClassID - {doc.ClassID}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Comments - {doc.Comments}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Company - {doc.Company}\n");
            Common.WriteToDebugWindow($"ActiveDocument.CompatibilityMode - {doc.CompatibilityMode}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Creator - {doc.Creator}\n");
            Common.WriteToDebugWindow($"ActiveDocument.DefaultFillStyle - {doc.DefaultFillStyle}\n");
            Common.WriteToDebugWindow($"ActiveDocument.DefaultGuideStyle - {doc.DefaultGuideStyle}\n");
            Common.WriteToDebugWindow($"ActiveDocument.DefaultLineStyle - {doc.DefaultLineStyle}\n");
            Common.WriteToDebugWindow($"ActiveDocument.DefaultSavePath - {doc.DefaultSavePath}\n");
            Common.WriteToDebugWindow($"ActiveDocument.DefaultStyle - {doc.DefaultStyle}\n");
            Common.WriteToDebugWindow($"ActiveDocument.DefaultTextStyle - {doc.DefaultTextStyle}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Description - {doc.Description}\n");
            Common.WriteToDebugWindow($"ActiveDocument.DefaultGuideStyle - {doc.DefaultGuideStyle}\n");
            Common.WriteToDebugWindow($"ActiveDocument.DocumentSheet.Name - {doc.DocumentSheet.Name}\n");
            Common.WriteToDebugWindow($"ActiveDocument.DynamicGridEnabled - {doc.DynamicGridEnabled}\n");
            Common.WriteToDebugWindow($"ActiveDocument.EditorCount - {doc.EditorCount}\n");
            Common.WriteToDebugWindow($"ActiveDocument.FooterCenter - {doc.FooterCenter}\n");
            Common.WriteToDebugWindow($"ActiveDocument.FooterLeft - {doc.FooterLeft}\n");
            Common.WriteToDebugWindow($"ActiveDocument.FooterMargin - {doc.FooterMargin}\n");
            Common.WriteToDebugWindow($"ActiveDocument.FooterRight - {doc.FooterRight}\n");
            Common.WriteToDebugWindow($"ActiveDocument.FullName - {doc.FullName}\n");
            Common.WriteToDebugWindow($"ActiveDocument.GestureFormatSheet.Name - {doc.GestureFormatSheet.Name}\n");
            Common.WriteToDebugWindow($"ActiveDocument.GlueEnabled - {doc.GlueEnabled}\n");
            Common.WriteToDebugWindow($"ActiveDocument.HeaderCenter - {doc.HeaderCenter}\n");
            Common.WriteToDebugWindow($"ActiveDocument.HeaderFooterColor - {doc.HeaderFooterColor}\n");

            // TODO(crhodes)
            // This throws exception.  Figure out why
            //Common.WriteToDebugWindow($"ActiveDocument.HeaderFooterFont - {doc.HeaderFooterFont}\n");

            var foo = doc.HeaderFooterFont;

            Common.WriteToDebugWindow($"ActiveDocument.HeaderLeft - {doc.HeaderLeft}\n");
            Common.WriteToDebugWindow($"ActiveDocument.HeaderMargin - {doc.HeaderMargin}\n");
            Common.WriteToDebugWindow($"ActiveDocument.HeaderRight - {doc.HeaderRight}\n");
            Common.WriteToDebugWindow($"ActiveDocument.HyperlinkBase - {doc.HyperlinkBase}\n");
            Common.WriteToDebugWindow($"ActiveDocument.ID - {doc.ID}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Index - {doc.Index}\n");
            Common.WriteToDebugWindow($"ActiveDocument.InPlace - {doc.InPlace}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Keywords - {doc.Keywords}\n");
            Common.WriteToDebugWindow($"ActiveDocument.LeftMargin - {doc.LeftMargin}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Manager - {doc.Manager}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Masters.Count - {doc.Masters.Count}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Pages.Count - {doc.Pages.Count}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Path - {doc.Path}\n");
            Common.WriteToDebugWindow($"ActiveDocument.PrintCenteredH - {doc.PrintCenteredH}\n");
            Common.WriteToDebugWindow($"ActiveDocument.PrintCenteredV - {doc.PrintCenteredV}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Printer - {doc.Printer}\n");
            Common.WriteToDebugWindow($"ActiveDocument.PrintFitOnPages - {doc.PrintFitOnPages}\n");
            Common.WriteToDebugWindow($"ActiveDocument.PrintLandscape - {doc.PrintLandscape}\n");
            Common.WriteToDebugWindow($"ActiveDocument.PrintPagesAcross - {doc.PrintPagesAcross}\n");
            Common.WriteToDebugWindow($"ActiveDocument.PrintPagesDown - {doc.PrintPagesDown}\n");
            Common.WriteToDebugWindow($"ActiveDocument.ProgID - {doc.ProgID}\n");
            Common.WriteToDebugWindow($"ActiveDocument.ReadOnly - {doc.ReadOnly}\n");
            Common.WriteToDebugWindow($"ActiveDocument.RemovePersonalInformation - {doc.RemovePersonalInformation}\n");
            Common.WriteToDebugWindow($"ActiveDocument.RightMargin - {doc.RightMargin}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Saved - {doc.Saved}\n");
            Common.WriteToDebugWindow($"ActiveDocument.SnapEnabled - {doc.SnapEnabled}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Stat - {doc.Stat}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Subject - {doc.Subject}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Template - {doc.Template}\n");
            Common.WriteToDebugWindow($"ActiveDocument.TimeCreated - {doc.TimeCreated}\n");
            Common.WriteToDebugWindow($"ActiveDocument.TimeEdited - {doc.TimeEdited}\n");
            Common.WriteToDebugWindow($"ActiveDocument.TimePrinted - {doc.TimePrinted}\n");
            Common.WriteToDebugWindow($"ActiveDocument.TimeSaved - {doc.TimeSaved}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Title - {doc.Title}\n");
            Common.WriteToDebugWindow($"ActiveDocument.TopMargin - {doc.TopMargin}\n");
            Common.WriteToDebugWindow($"ActiveDocument.UndoEnabled - {doc.UndoEnabled}\n");
            Common.WriteToDebugWindow($"ActiveDocument.Version - {doc.Version}\n");
        }

        public static void DisplayLayer(string layerName, bool show)
        {
            Common.WriteToDebugWindow(string.Format("{0}(layer:{1} show:{2})",
                MethodBase.GetCurrentMethod().Name, layerName, show.ToString()));

            foreach (MSVisio.Page page in Common.VisioApplication.ActiveDocument.Pages)
            {
                Visio_Page.DisplayLayer(page, layerName, show);
            }
        }

        public static void DisplayPageNames()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Document doc = app.ActiveDocument;

            try
            {
                foreach (MSVisio.Page page in doc.Pages)
                {
                    Common.WriteToDebugWindow(string.Format("Page {0} Name:>{1:30}< NameU:>{2:30}<",
                        page.Name.Equals(page.NameU) == true ? " " : "?",
                        page.Name, page.NameU));
                }
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force:true);
            }
        }

        public static List<MSVisio.Shape> GetNavigationLinks()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            List<MSVisio.Shape> navLinks = new List<MSVisio.Shape>();

            MSVisio.Page linkPage = Common.VisioApplication.ActiveDocument.Pages["Navigation Links"];

            foreach (MSVisio.Shape shape in linkPage.Shapes)
            {
                navLinks.Add(shape);
            }

            return navLinks;
        }

        public static void MovePages(string targetDocument)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;
            MSVisio.Document doc = app.ActiveDocument;
            MSVisio.Page page = app.ActivePage;

            foreach (MSVisio.Shape shape in page.Shapes)
            {
                Common.WriteToDebugWindow(string.Format("Name: {0}  Text: {1}", shape.Name, shape.Text));
                try
                {
                    Visio_Page.MovePage(doc.Pages[shape.Text], targetDocument);
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow(ex.ToString(), force:true);
                }
            }
        }

        public static void PrintPages()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            try
            {
                MSVisio.Application app = Common.VisioApplication;
                MSVisio.Document doc = app.ActiveDocument;
                MSVisio.Page page = app.ActivePage;

                foreach (MSVisio.Shape shape in page.Shapes)
                {
                    Common.WriteToDebugWindow(string.Format("Name: {0}  Text: >{1}<", shape.Name, shape.Text));

                    //var bar = shape.Hyperlink;
                    //var hyperLinks = shape.Hyperlinks;

                    //foreach (MSVisio.Hyperlink hyperlink in hyperLinks)
                    //{
                    //    var foo = hyperlink.Address;
                    //}

                    if (shape.Hyperlinks.Count > 0)
                    {
                        if (shape.Hyperlink.SubAddress.Length > 0)
                        {
                            Common.WriteToDebugWindow(string.Format("Hyperlink: >{0}<", shape.Hyperlink.SubAddress));
                            doc.Pages[shape.Hyperlink.SubAddress].Print();
                        }
                    }
                }
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

            foreach (MSVisio.Page page in doc.Pages)
            {
                Visio_Page.DeleteLayers(page);
            }
        }

        public static void RenamePages(string searchExpression, string replacementExpression,
            RegexOptions regexOptions = RegexOptions.None)
        {
            VNC.Log.Trace("", Common.LOG_CATEGORY, 0);

            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;
            MSVisio.Document doc = app.ActiveDocument;
            MSVisio.Page page = app.ActivePage;

            Regex regex = new Regex(searchExpression, regexOptions);

            foreach (MSVisio.Shape shape in page.Shapes)
            {
                Common.WriteToDebugWindow(string.Format("Name: {0}  Text: {1}", shape.Name, shape.Text));
                try
                {
                    string newPageName = regex.Replace(doc.Pages[shape.Text].Name, replacementExpression);

                    doc.Pages[shape.Text].Name = newPageName;
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow(ex.ToString(), force:true);
                }
            }
        }

        public static void SavePages()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;
            MSVisio.Document doc = app.ActiveDocument;
            MSVisio.Page page = app.ActivePage;

            foreach (MSVisio.Shape shape in page.Shapes)
            {
                Common.WriteToDebugWindow(string.Format("Name: {0}  Text: {1}", shape.Name, shape.Text));
                try
                {
                    Visio_Page.SavePage(doc.Pages[shape.Text]);
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow(ex.ToString(), force:true);
                }
            }
        }

        public static void SortAllPages()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Document doc = app.ActiveDocument;

            System.Collections.SortedList sortedPages = new System.Collections.SortedList();
            //SortedList<string, string> sortedPages = new SortedList<string, string>();
            int index = 0;
            bool hasTOCPage = false;

            Common.WriteToDebugWindow(string.Format("Document({0})", doc.Name));

            try
            {
                foreach (MSVisio.Page page in doc.Pages)
                {
                    Common.WriteToDebugWindow(string.Format("Page({0} IsBackground{1})", page.NameU, page.Background));

                    if (!page.NameU.Equals("Table of Contents"))
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
                    Common.WriteToDebugWindow(string.Format("Moving Page({0})", sortedPages.GetByIndex(i)));
                    doc.Pages.ItemU[sortedPages.GetKey(i)].Index = (short)(i + offset);
                    //Application.ActiveDocument.Pages.ItemU("Page-2").Index = 3
                }
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force:true);
            }
        }

        public static void SyncPageNames()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Document doc = app.ActiveDocument;

            List<MSVisio.Page> pagesToUpdate = new List<MSVisio.Page>();

            Common.WriteToDebugWindow(string.Format("Document({0})", doc.Name));

            foreach (MSVisio.Page page in doc.Pages)
            {
                if (!page.Name.Equals(page.NameU))
                {
                    pagesToUpdate.Add(page);
                }
            }

            foreach (MSVisio.Page page in pagesToUpdate)
            {
                Visio_Page.SyncPageNames(page);
            }
        }

        public static void UpdatePageNameShapes()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Document doc = app.ActiveDocument;

            foreach (MSVisio.Page page in doc.Pages)
            {
                Visio_Page.UpdatePageNameShapes(page);
            }
        }

        #endregion

        #region Private Methods

        public static void AddArchitectureBasePages()
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            MSVisio.Application app = Common.VisioApplication;

            MSVisio.Page page = null;
            MSVisio.Document archStencil = null;
            MSVisio.Master archMaster = null;

            int undoScope = Common.VisioApplication.BeginUndoScope("AddArchitectureBasePages");

            var architecturePages = new string[] {
                "Clean Architecture 1",
                "Clean Architecture 2",
                "Clean Architecture 3",
                "Clean Architecture 4" };

            // Delete any existing Architecture Base Pages

            foreach (var existingPage in architecturePages)
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
                    //Common.WriteToDebugWindow(ex.ToString(), force:true);
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
                app.Documents.OpenEx(archStencilName, (short)MSVisio.VisOpenSaveArgs.visOpenRO + (short)MSVisio.VisOpenSaveArgs.visOpenDocked);

                archStencil = app.Documents[archStencilName];
            }

            // Assume 11 x 8.5 Landscape page

            try
            {
                page = Common.VisioApplication.ActiveDocument.Pages.Add();
                archMaster = archStencil.Masters["Clean Arch 1 - Page Base"];

                page.NameU = "Clean Architecture 1";
                page.Background = 1;
                page.Drop(archMaster, 5.5, 4.325);

                page = Common.VisioApplication.ActiveDocument.Pages.Add();
                archMaster = archStencil.Masters["Clean Arch 2 - Page Base"];

                page.NameU = "Clean Architecture 2";
                page.Background = 1;
                page.Drop(archMaster, 5.5, 4.325);

                page = Common.VisioApplication.ActiveDocument.Pages.Add();
                archMaster = archStencil.Masters["Clean Arch 3 - Page Base"];

                page.NameU = "Clean Architecture 3";
                page.Background = 1;
                page.Drop(archMaster, 5.5, 4.325);

                page = Common.VisioApplication.ActiveDocument.Pages.Add();
                archMaster = archStencil.Masters["Clean Arch 4 - Page Base"];

                page.NameU = "Clean Architecture 4";
                page.Background = 1;
                page.Drop(archMaster, 5.5, 4.325);
            }
            catch (Exception ex)
            {

            }

            Common.VisioApplication.EndUndoScope(undoScope, true);
        }

        public static void CreatePluralSightCourseFileFromShape(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] array)
        {
            int i = 0;

            MSVisio.Page currentPage = app.ActivePage;
            MSVisio.Document currentDocument = app.ActiveDocument;
            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
            string shapePageName = "COURSENAME";
            string shapeAuthor = "AUTHOR";

            try
            {
                if (activeShape.CellExistsU["Prop.PageName", 0] != 0)
                {
                    shapePageName = activeShape.CellsU["Prop.PageName"].ResultStrU[MSVisio.VisUnitCodes.visUnitsString];
                }

                if (activeShape.CellExistsU["Prop.Author", 0] != 0)
                {
                    shapeAuthor = activeShape.CellsU["Prop.Author"].ResultStrU[MSVisio.VisUnitCodes.visUnitsString];
                }

                string templateName = "CHR Notes - PluralSight Course - Subject - Author.vstx";
                app.Documents.AddEx(templateName, MSVisio.VisMeasurementSystem.visMSDefault, 0, 0);

                app.ActivePage.Drop(activeShape, (short)5.5, (short)4.25);

                string fileName = $"CHR Notes - PluralSight Course - {shapePageName} - {shapeAuthor}.vsdx";

                SaveFileDialog saveFileDiaglog = new SaveFileDialog();

                saveFileDiaglog.FileName = fileName;
                saveFileDiaglog.InitialDirectory = @"B:\CHR Notes\PluralSight";

                DialogResult result = saveFileDiaglog.ShowDialog();

                fileName = saveFileDiaglog.FileName;

                app.ActiveDocument.SaveAsEx(fileName, (short)MSVisio.VisOpenSaveArgs.visSaveAsWS);
            }
            catch (Exception ex)
            {

            }
        }

        private static void AddPageLinkToTOCPage(MSVisio.Page pageTOC, MSVisio.Page page, double xLoc, double yLoc)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            int undoScope = Common.VisioApplication.BeginUndoScope("AddPageLinkToTOCPage");

            MSVisio.Shape pageLinkShape = pageTOC.DrawRectangle(xLoc, yLoc, xLoc + cTOC_Link_Width, yLoc + cTOC_Link_Height);

            pageLinkShape.TextStyle = "Normal";
            pageLinkShape.LineStyle = "Text Only";
            pageLinkShape.FillStyle = "Text Only";
            pageLinkShape.Characters.Begin = 0;
            pageLinkShape.Characters.End = 0;
            pageLinkShape.Text = page.NameU;
            pageLinkShape.Characters.set_CharProps((short)MSVisio.VisCellIndices.visCharacterSize, cTOC_Link_FontSize);

            Visio_Shape.AddHyperLink(pageLinkShape, page.NameU);

            Common.VisioApplication.EndUndoScope(undoScope, true);
        }

        private static void AddTOCLinkToPage(MSVisio.Page page)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            if (page.Background != 0)
            {
                // Skip background pages
                return;
            }

            int undoScope = Common.VisioApplication.BeginUndoScope("AddTOCLinkToPage");

            // Clear out any existing link.

            foreach (MSVisio.Shape shape in page.Shapes)
            {
                if (shape.Text == "Table of Contents" || shape.Name == "TOCLink")
                {
                    shape.Delete();
                }
            }

            MSVisio.Shape tocShape = page.DrawRectangle(
                cTOC_Page_Initial_xLoc, cTOC_PageLink_Initial_yLoc,
                cTOC_Page_Initial_xLoc + cTOC_PageLink_Width, cTOC_PageLink_Initial_yLoc + cTOC_PageLink_Height);

            tocShape.Name = "TOCLink";

            tocShape.Text = "Table of Contents";
            tocShape.TextStyle = "Normal";
            tocShape.LineStyle = "Text Only";
            tocShape.FillStyle = "Text Only";
            tocShape.Characters.Begin = 0;
            tocShape.Characters.End = 0;
            tocShape.Characters.set_CharProps((short)MSVisio.VisCellIndices.visCharacterSize, 6);

            MSVisio.Hyperlink hlink = tocShape.Hyperlinks.Add();
            // hlink.Name = "do we need a name?";
            hlink.SubAddress = "Table of Contents";

            Common.VisioApplication.EndUndoScope(undoScope, true);
        }

        private static void ClearPage(MSVisio.Page page)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            try
            {
                for (int i = page.Shapes.Count - 1; i >= 0; i--)
                {
                    page.Shapes[i].Delete();
                }
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force:true);
            }

            System.Diagnostics.Debug.WriteLine(string.Format("Shapes on Page: {0}", page.Shapes.Count));

        }
        #endregion
    }
}
