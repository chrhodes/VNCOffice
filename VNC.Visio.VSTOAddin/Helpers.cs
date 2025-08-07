using System;
using System.Reflection;
using System.Windows;

using VNC.Core;

using MSVisio = Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn
{
    public class Helpers
    {
        public static void Add_ActionSection_Row(
            MSVisio.Shape shape, 
            string rowName,
            string action,
            string menu,
            string tagName = "",
            string buttonFace = "",
            string sortKey = "",
            string isChecked = "0",
            string isDisabled = "0",
            string isReadOnly = "0",
            string isInvisible = "0",
            string beginGroup = "0",
            string flyoutChild = "0")
        {
            //result = shape.AddRow((short)MSVisio.VisSectionIndices.visSectionAction, (short)MSVisio.VisRowIndices.visRowLast, (short)MSVisio.VisRowTags.visTagDefault);
            // TODO(crhodes):
            // Determine what this does if row already exists.

            try
            {
                var rowNumber = shape.AddNamedRow((short)MSVisio.VisSectionIndices.visSectionAction, rowName, (short)MSVisio.VisRowTags.visTagDefault);

                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionAction].FormulaU = action;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionMenu].FormulaU = menu.WrapInDblQuotes();
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionTagName].FormulaU = tagName;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionButtonFace].FormulaU = buttonFace;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionSortKey].FormulaU = sortKey;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionChecked].FormulaU = isChecked;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionDisabled].FormulaU = isDisabled;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionReadOnly].FormulaU = isReadOnly;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionInvisible].FormulaU = isInvisible;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionBeginGroup].FormulaU = beginGroup;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionFlyoutChild].FormulaU = flyoutChild;
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force: true);
            }
        }

        public static void Add_HyperlinkSection_Row(
            MSVisio.Shape shape,
            string rowName,
            string description,
            string address,
            string subAddress,
            string extraInfo = "",
            string frame = "",
            string sortKey = "",
            string newWindow = "0",
            string default1 = "0",
            string invisible = "0")
        {
            //result = shape.AddRow((short)MSVisio.VisSectionIndices.visSectionAction, (short)MSVisio.VisRowIndices.visRowLast, (short)MSVisio.VisRowTags.visTagDefault);
            // TODO(crhodes):
            // Determine what this does if row already exists.

            var rowNumber = shape.AddNamedRow(
                (short)MSVisio.VisSectionIndices.visSectionHyperlink, 
                rowName, 
                (short)MSVisio.VisRowTags.visTagDefault);

            try
            {
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visHLinkDescription].FormulaU = description;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visHLinkAddress].FormulaU = address;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visHLinkSubAddress].FormulaU = subAddress;  // Wrapping in doubleqoutes would break entering formulas
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visHLinkExtraInfo].FormulaU = extraInfo;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visHLinkExtraInfo].FormulaU = frame;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visHLinkSortKey].FormulaU = sortKey;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visHLinkNewWin].FormulaU = newWindow;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visHLinkDefault].FormulaU = default1;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionHyperlink,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visHLinkInvisible].FormulaU = invisible;

            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force: true);
            }

        }
        /// <summary>
        /// Add a Prop (ShapeData) section to a ShapeSheet
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="rowName"></param>
        /// <param name="label"></param>
        /// <param name="type"></param>
        /// <param name="format">Values must be placed in Quotes ("") if strings</param>
        /// <param name="value">Values must be placed in Quotes ("") if strings</param>
        /// <param name="prompt"></param>
        /// <param name="sortKey"></param>
        /// <param name="invisible"></param>
        /// <param name="ask"></param>
        /// <param name="langID"></param>
        /// <param name="calendar"></param>
        public static void Add_Prop_Row(MSVisio.Shape shape,
            string rowName,
            string label, short type, string format, string value,
            string prompt = null, string sortKey = null,
            string invisible = null, string ask = null, string langID = null, string calendar = null)
        {
            Validate_Prop_SectionExists(shape);

            try
            {
                // Add the Row

                short rowNumber = shape.AddNamedRow(
                    (short)MSVisio.VisSectionIndices.visSectionProp,
                    rowName,
                    (short)MSVisio.VisRowTags.visTagDefault);

                // And the important cells: Label, Type, Value

                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionProp,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visCustPropsLabel].FormulaU = label.WrapInDblQuotes();

                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionProp,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visCustPropsType].FormulaU = type.ToString();

                if (format != null)
                {
                    shape.CellsSRC[
                        (short)MSVisio.VisSectionIndices.visSectionProp,
                        rowNumber,
                        (short)MSVisio.VisCellIndices.visCustPropsFormat].FormulaU = format.WrapInDblQuotes();    // Is this ever wrong?
                }

                var v1 = value;
                var v2 = value;

                if (value.Contains("\""))
                {
                    value = value.Replace("\"", "\"\"");
                }

                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionProp,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visCustPropsValue].FormulaU = value.WrapInDblQuotes();    // Is this ever wrong?;

                // And any optional cells

                if (!String.IsNullOrEmpty(prompt))
                {
                    shape.CellsSRC[
                       (short)MSVisio.VisSectionIndices.visSectionProp,
                       rowNumber,
                       (short)MSVisio.VisCellIndices.visCustPropsPrompt].FormulaU = prompt.WrapInDblQuotes();
                }

                //if (null != prompt)
                //{
                //    shape.CellsSRC[
                //       (short)MSVisio.VisSectionIndices.visSectionProp,
                //       rowNumber,
                //       (short)MSVisio.VisCellIndices.visCustPropsPrompt].FormulaU = prompt.WrapInDblQuotes();
                //}

                if (null != sortKey)
                {
                    shape.CellsSRC[
                        (short)MSVisio.VisSectionIndices.visSectionProp,
                        rowNumber,
                        (short)MSVisio.VisCellIndices.visCustPropsSortKey].FormulaU = sortKey.WrapInDblQuotes();
                }

                // TODO(crhodes):
                // Add support for remaining optional arguments

            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force: true);
            }
        }

        public static void Add_ShapeDataSection_Row(
            MSVisio.Shape shape, 
            string rowName,
            string action,
            string menu,
            string tagName = "",
            string buttonFace = "",
            string sortKey = "",
            string isChecked = "0",
            string isDisabled = "0",
            string isReadOnly = "0",
            string isInvisible = "0",
            string beginGroup = "0",
            string flyoutChild = "0")
        {
            var rowNumber = shape.AddNamedRow((short)MSVisio.VisSectionIndices.visSectionAction, rowName, (short)MSVisio.VisRowTags.visTagDefault);

            try
            {
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionAction].FormulaU = action;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionMenu].FormulaU = menu;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionTagName].FormulaU = tagName;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionButtonFace].FormulaU = buttonFace;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionSortKey].FormulaU = sortKey;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionChecked].FormulaU = isChecked;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionDisabled].FormulaU = isDisabled;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionReadOnly].FormulaU = isReadOnly;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionInvisible].FormulaU = isInvisible;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionBeginGroup].FormulaU = beginGroup;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionAction,
                    rowNumber,
                    (short)MSVisio.VisCellIndices.visActionFlyoutChild].FormulaU = flyoutChild;
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force: true);
            }
        }

        public static void Add_User_Row(MSVisio.Shape shape,
            string rowName, string value,
            string prompt = "")
        {
            Validate_User_SectionExists(shape);

            try
            {
                short rowNumber = shape.AddNamedRow(
                    (short)MSVisio.VisSectionIndices.visSectionUser,
                    rowName,
                    (short)MSVisio.VisRowTags.visTagDefault);

                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionUser,
                    (short)(MSVisio.VisRowIndices.visRowControl + rowNumber),
                    (short)MSVisio.VisCellIndices.visUserValue].FormulaU = value;

                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionUser,
                    (short)(MSVisio.VisRowIndices.visRowControl + rowNumber),
                    (short)MSVisio.VisCellIndices.visUserPrompt].FormulaU = string.Format("\"{0}\"", prompt);
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force: true);
            }
        }

        public static void Delete_Section_Row(
            MSVisio.Shape shape,
            MSVisio.VisSectionIndices sectionIndex,
            string sectionName,
            string rowName)
        {
            try
            {
                Validate_Prop_SectionExists(shape);

                short rowIndex = -1;

                if (shape.CellExistsU[$"{sectionName}.{rowName}", 0] != 0)
                {
                    rowIndex = shape.CellsRowIndex[$"{sectionName}.{rowName}"];
                    shape.DeleteRow((short)sectionIndex, rowIndex);
                }
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force: true);
            }
        }

        public static short GetVisPropType(string value)
        {
            short visPropType = 0;

            switch (value)
            {
                case "VisCellVals.visPropTypeBool":
                case "3":
                    visPropType = (short)MSVisio.VisCellVals.visPropTypeBool;
                    break;

                case "VisCellVals.visPropTypeCurrency":
                case "7":
                    visPropType = (short)MSVisio.VisCellVals.visPropTypeCurrency;
                    break;

                case "VisCellVals.visPropTypeDate":
                case "5":
                    visPropType = (short)MSVisio.VisCellVals.visPropTypeDate;
                    break;

                case "VisCellVals.visPropTypeDuration":
                case "6":
                    visPropType = (short)MSVisio.VisCellVals.visPropTypeDuration;
                    break;

                case "VisCellVals.visPropTypeListFix":
                case "1":
                    visPropType = (short)MSVisio.VisCellVals.visPropTypeListFix;
                    break;

                case "VisCellVals.visPropTypeListVar":
                case "4":
                    visPropType = (short)MSVisio.VisCellVals.visPropTypeListVar;
                    break;

                case "VisCellVals.visPropTypeNumber":
                case "2":
                    visPropType = (short)MSVisio.VisCellVals.visPropTypeNumber;
                    break;

                case "VisCellVals.visPropTypeString":
                case "0":
                    visPropType = (short)MSVisio.VisCellVals.visPropTypeString;
                    break;

                default:
                    Common.WriteToWatchWindow(string.Format("Unrecognized VisPropType >{0}<", value));

                    break;
            }

            return visPropType;
        }

        public static string GetShapePropertyAsString(MSVisio.Shape activeShape, string property)
        {
            string propertyName = "Prop." + property;
            string result = "";

            if (activeShape.CellExistsU[propertyName, 0] != 0)
            {
                result = activeShape.CellsU[propertyName].ResultStrU[MSVisio.VisUnitCodes.visUnitsString];
            }

            return result;
        }

        public static bool LoadStencil(MSVisio.Application app, string stencilName)
        {
            bool result = false;

            try
            {
                var foo = app.Documents[stencilName];
                result = true;
            }
            catch (Exception)
            {
                // Stencil may not be open.  Try opening it

                try
                {
                    app.Documents.OpenEx(stencilName, (short)MSVisio.VisOpenSaveArgs.visOpenRO + (short)MSVisio.VisOpenSaveArgs.visOpenDocked);
                    result = true;
                }
                catch (Exception)
                {
                    MessageBox.Show($"Cannot locate or open {stencilName}, aborting.");
                }
            }

            return result;
        }

        public static void Populate_Actions_Section(
            MSVisio.Shape shape, 
            string actionName, 
            string action, 
            string menu, 
            string tagName, 
            string buttonFace, 
            string sortKey, 
            string isChecked, 
            string isDisabled, 
            string isReadOnly, 
            string isInvisible, 
            string beginGroup, 
            string flyoutChild)
        {
            Add_ActionSection_Row(shape,
                actionName,
                action,
                menu,
                tagName,
                buttonFace, sortKey, isChecked, isDisabled, isReadOnly, isInvisible, beginGroup, flyoutChild);
        }

        public static void Populate_Hyperlinks_Section(
            MSVisio.Shape shape, 
            string rowName, 
            string description, 
            string address, 
            string subAddress, 
            string extraInfo, 
            string frame, 
            string sortKey, 
            string newWindow, 
            string default1, 
            string invisible)
        {
            Add_HyperlinkSection_Row(shape,
                rowName,
                description,
                address,
                subAddress,
                extraInfo, frame, sortKey, newWindow, default1, invisible);
        }

        public static void Populate_Controls_Section(
            MSVisio.Shape shape,
            string X, string Y,
            string XDynamics, string YDynamics,
            string XBehavior, string YBehavior,
            string CanGlue, string Tip)
        {
            // There can be more than one Controls Row so need to think through how to handle existing rows.

            Validate_Controls_SectionExists(shape);

            short newRow = shape.AddRow(
                (short)MSVisio.VisSectionIndices.visSectionControls,
                (short)MSVisio.VisRowIndices.visRowControl,
                (short)MSVisio.VisRowTags.visTagDefault);

            try
            {
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionControls,
                    (short)MSVisio.VisRowIndices.visRowControl + 0,
                    (short)MSVisio.VisCellIndices.visCtlX].FormulaU = X;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionControls,
                    (short)MSVisio.VisRowIndices.visRowControl + 0,
                    (short)MSVisio.VisCellIndices.visCtlY].FormulaU = Y;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionControls,
                    (short)MSVisio.VisRowIndices.visRowControl + 0,
                    (short)MSVisio.VisCellIndices.visCtlXDyn].FormulaU = XDynamics;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionControls,
                    (short)MSVisio.VisRowIndices.visRowControl + 0,
                    (short)MSVisio.VisCellIndices.visCtlYDyn].FormulaU = YDynamics;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionControls,
                    (short)MSVisio.VisRowIndices.visRowControl + 0,
                    (short)MSVisio.VisCellIndices.visCtlXCon].FormulaU = XBehavior;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionControls,
                    (short)MSVisio.VisRowIndices.visRowControl + 0,
                    (short)MSVisio.VisCellIndices.visCtlYCon].FormulaU = YBehavior;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionControls,
                    (short)MSVisio.VisRowIndices.visRowControl + 0,
                    (short)MSVisio.VisCellIndices.visCtlGlue].FormulaU = CanGlue;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionControls,
                    (short)MSVisio.VisRowIndices.visRowControl + 0,
                    (short)MSVisio.VisCellIndices.visCtlTip].FormulaU = string.Format("\"{0}\"", Tip);
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force: true);
            }
        }

        public static void Populate_Controls_Section(
            MSVisio.Shape shape, 
            string rowName,
            string X, string Y,
            string XDynamics, string YDynamics,
            string XBehavior, string YBehavior,
            string CanGlue, 
            string Tip)
        {
            // There can be more than one Controls Row so need to think through how to handle existing rows.

            Validate_Controls_SectionExists(shape);

            short newRow = shape.AddNamedRow(
                (short)MSVisio.VisSectionIndices.visSectionControls,
                rowName,
                (short)MSVisio.VisRowTags.visTagDefault);

            try
            {
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionControls,
                    newRow,
                    (short)MSVisio.VisCellIndices.visCtlX].FormulaU = X;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionControls,
                    newRow,
                    (short)MSVisio.VisCellIndices.visCtlY].FormulaU = Y;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionControls,
                    newRow,
                    (short)MSVisio.VisCellIndices.visCtlXDyn].FormulaU = XDynamics;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionControls,
                    newRow,
                    (short)MSVisio.VisCellIndices.visCtlYDyn].FormulaU = YDynamics;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionControls,
                    newRow,
                    (short)MSVisio.VisCellIndices.visCtlXCon].FormulaU = XBehavior;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionControls,
                    newRow,
                    (short)MSVisio.VisCellIndices.visCtlYCon].FormulaU = YBehavior;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionControls,
                    newRow,
                    (short)MSVisio.VisCellIndices.visCtlGlue].FormulaU = CanGlue;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionControls,
                    newRow,
                    (short)MSVisio.VisCellIndices.visCtlTip].FormulaU = string.Format("\"{0}\"", Tip);
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force: true);
            }
        }

        public static Boolean RowExists(
            MSVisio.Shape shape,
            MSVisio.VisSectionIndices sectionIndex,
            MSVisio.VisRowIndices rowIndex,
            MSVisio.VisExistsFlags visExistsFlags = MSVisio.VisExistsFlags.visExistsAnywhere)
        {
            if (0 == shape.RowExists[(short)sectionIndex, (short)rowIndex, (short)visExistsFlags])
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public static string SafeFileName(string fileName)
        {
            fileName = fileName.Replace("/", "-");
            fileName = fileName.Replace(@"\", "-");
            fileName = fileName.Replace("[", "");
            fileName = fileName.Replace("]", "");
            //fileName = fileName.Replace(" ", "");
            fileName = fileName.Replace(":", "-");

            return fileName;
        }

        public static string SafePageName(string pageName)
        {
            pageName = pageName.Replace("/", "-");
            pageName = pageName.Replace(@"\", "-");
            pageName = pageName.Replace("[", "");
            pageName = pageName.Replace("]", "");
            //pageName = pageName.Replace(" ", "");
            pageName = pageName.Replace("\n", " ");
            pageName = pageName.Replace(":", "-");

            return pageName;
        }
        public static void Set_FillFormat_SectionOld(MSVisio.Shape shape,
            string fillForegnd = null, string fillForegndTrans = null,
            string fillBkgnd = null, string fillBkgndTrans = null, string fillPattern = null,
            string shdwForegnd = null, string shdwForegndTrans = null,
            string shdwPattern = null, string shapeShdwOffsetX = null, string shapeShdwOffsetY = null,
            string shapeShdwType = null, string shapeShdwObliqueAngle = null, string shapeShdwScaleFactor = null,
            string shapeShdwBlur = null, string shapeShdwShow = null)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            // This Section always exists, so just set values

            // Everything defaults to null and is in the likely order of most often changed.
            // If null, skip setting value.

            try
            {
                Set_RowFill_Cell(shape, MSVisio.VisCellIndices.visFillForegnd, fillForegnd);
                Set_RowFill_Cell(shape, MSVisio.VisCellIndices.visFillForegndTrans, fillForegndTrans);

                Set_RowFill_Cell(shape, MSVisio.VisCellIndices.visFillBkgnd, fillBkgnd);
                Set_RowFill_Cell(shape, MSVisio.VisCellIndices.visFillBkgndTrans, fillBkgndTrans);
                Set_RowFill_Cell(shape, MSVisio.VisCellIndices.visFillPattern, fillPattern);

                Set_RowFill_Cell(shape, MSVisio.VisCellIndices.visFillShdwForegnd, shdwForegnd);
                Set_RowFill_Cell(shape, MSVisio.VisCellIndices.visFillShdwForegndTrans, shdwForegndTrans);

                Set_RowFill_Cell(shape, MSVisio.VisCellIndices.visFillShdwPattern, shdwPattern);
                Set_RowFill_Cell(shape, MSVisio.VisCellIndices.visFillShdwOffsetX, shapeShdwOffsetX);
                Set_RowFill_Cell(shape, MSVisio.VisCellIndices.visFillShdwOffsetY, shapeShdwOffsetY);

                Set_RowFill_Cell(shape, MSVisio.VisCellIndices.visFillShdwType, shapeShdwType);
                Set_RowFill_Cell(shape, MSVisio.VisCellIndices.visFillShdwObliqueAngle, shapeShdwObliqueAngle);
                Set_RowFill_Cell(shape, MSVisio.VisCellIndices.visFillShdwScaleFactor, shapeShdwScaleFactor);
                Set_RowFill_Cell(shape, MSVisio.VisCellIndices.visFillShdwBlur, shapeShdwBlur);
                Set_RowFill_Cell(shape, MSVisio.VisCellIndices.visFillShdwShow, shapeShdwShow);
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force: true);
            }
        }

        public static void Set_RowFill_Cell(
                                                                                                                            MSVisio.Shape shape,
            MSVisio.VisCellIndices cellIndex,
            string value)
        {
            if (value != null)
            {
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionObject,
                    (short)MSVisio.VisRowIndices.visRowFill,
                    (short)cellIndex].FormulaU = value;
            }
        }

        public static void Set_ShapeTransform_Section(MSVisio.Shape shape,
                        string width = null, string height = null, string pinX = null, string pinY = null,
                string flipX = null, string flipY = null, string locPinX = null, string locPinY = null,
                string angle = null, string resizeMode = null)
        {
            Common.WriteToDebugWindow($"{MethodBase.GetCurrentMethod().Name}()");

            // This Section always exists, so just set values

            // Everything defaults to null and is in the likely order of most often changed.
            // If null, skip setting value.

            try
            {
                Set_RowXFormOut_Cell(shape, MSVisio.VisCellIndices.visXFormWidth, width);
                Set_RowXFormOut_Cell(shape, MSVisio.VisCellIndices.visXFormHeight, height);
                Set_RowXFormOut_Cell(shape, MSVisio.VisCellIndices.visXFormPinX, pinX);
                Set_RowXFormOut_Cell(shape, MSVisio.VisCellIndices.visXFormPinY, pinY);
                Set_RowXFormOut_Cell(shape, MSVisio.VisCellIndices.visXFormFlipX, flipX);
                Set_RowXFormOut_Cell(shape, MSVisio.VisCellIndices.visXFormFlipY, flipY);
                Set_RowXFormOut_Cell(shape, MSVisio.VisCellIndices.visXFormLocPinX, locPinX);
                Set_RowXFormOut_Cell(shape, MSVisio.VisCellIndices.visXFormLocPinY, locPinY);
                Set_RowXFormOut_Cell(shape, MSVisio.VisCellIndices.visXFormAngle, angle);
                Set_RowXFormOut_Cell(shape, MSVisio.VisCellIndices.visXFormResizeMode, resizeMode);
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force: true);
            }
        }

        public static void Set_TextXForm_Section(MSVisio.Shape shape,
            string Width, string Height,
            string PinX, string PinY,
            string LocPinX, string LocPinY,
            string Angle)
        {
            Validate_TextXForm_SectionExists(shape);

            try
            {
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionObject,
                    (short)MSVisio.VisRowIndices.visRowTextXForm,
                    (short)MSVisio.VisCellIndices.visXFormWidth].FormulaU = Width;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionObject,
                    (short)MSVisio.VisRowIndices.visRowTextXForm,
                    (short)MSVisio.VisCellIndices.visXFormHeight].FormulaU = Height;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionObject,
                    (short)MSVisio.VisRowIndices.visRowTextXForm,
                    (short)MSVisio.VisCellIndices.visXFormPinX].FormulaU = PinX;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionObject,
                    (short)MSVisio.VisRowIndices.visRowTextXForm,
                    (short)MSVisio.VisCellIndices.visXFormPinY].FormulaU = PinY;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionObject,
                    (short)MSVisio.VisRowIndices.visRowTextXForm,
                    (short)MSVisio.VisCellIndices.visXFormLocPinX].FormulaU = LocPinX;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionObject,
                    (short)MSVisio.VisRowIndices.visRowTextXForm,
                    (short)MSVisio.VisCellIndices.visXFormLocPinY].FormulaU = LocPinY;
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionObject,
                    (short)MSVisio.VisRowIndices.visRowTextXForm,
                    (short)MSVisio.VisCellIndices.visXFormAngle].FormulaU = Angle;
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString(), force: true);
            }
        }

        public static void Validate_Action_SectionExists(MSVisio.Shape shape)
        {
            if (0 == shape.SectionExists[(short)MSVisio.VisSectionIndices.visSectionAction, 0])
            {
                try
                {
                    var result = shape.AddSection((short)MSVisio.VisSectionIndices.visSectionAction);
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow(ex.ToString(), force: true);
                }
            }
        }

        public static void Validate_Controls_SectionExists(MSVisio.Shape shape)
        {
            if (0 == shape.SectionExists[(short)MSVisio.VisSectionIndices.visSectionControls, 0])
            {
                try
                {
                    var result = shape.AddSection((short)MSVisio.VisSectionIndices.visSectionControls);
                    //result = shape.AddRow(
                    //    (short)MSVisio.VisSectionIndices.visSectionControls, 
                    //    (short)MSVisio.VisRowIndices.visRowControl, 
                    //    (short)MSVisio.VisRowTags.visTagDefault);
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow(ex.ToString(), force: true);
                }
            }
        }

        public static void Validate_HyperLink_SectionExists(MSVisio.Shape shape)
        {
            // NB. Shape Data = visSectionProp

            if (0 == shape.SectionExists[(short)MSVisio.VisSectionIndices.visSectionHyperlink, 0])
            {
                try
                {
                    var result = shape.AddSection((short)MSVisio.VisSectionIndices.visSectionHyperlink);
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow(ex.ToString(), force: true);
                }
            }
        }

        public static void Validate_Prop_SectionExists(MSVisio.Shape shape)
        {
            // NB. Shape Data = visSectionProp

            if (0 == shape.SectionExists[(short)MSVisio.VisSectionIndices.visSectionProp, 0])
            {
                try
                {
                    var result = shape.AddSection((short)MSVisio.VisSectionIndices.visSectionProp);
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow(ex.ToString(), force: true);
                }
            }
        }

        public static void Validate_TextField_SectionExists(MSVisio.Shape shape)
        {
            if (0 == shape.RowExists[
                (short)MSVisio.VisSectionIndices.visSectionTextField,
                (short)MSVisio.VisRowIndices.visRowText,
                (short)MSVisio.VisExistsFlags.visExistsAnywhere])
            {
                try
                {
                    shape.AddRow(
                        (short)MSVisio.VisSectionIndices.visSectionTextField,
                        (short)MSVisio.VisRowIndices.visRowText,
                        (short)MSVisio.VisRowTags.visTagDefault);
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow(ex.ToString(), force: true);
                }
            }
        }

        public static void Validate_TextXForm_SectionExists(MSVisio.Shape shape)
        {
            // TextXForm exists as a row in the SectionObject!

            if (0 == shape.RowExists[
                (short)MSVisio.VisSectionIndices.visSectionObject,
                (short)MSVisio.VisRowIndices.visRowTextXForm,
                (short)MSVisio.VisExistsFlags.visExistsAnywhere])
            {
                try
                {
                    shape.AddRow(
                        (short)MSVisio.VisSectionIndices.visSectionObject,
                        (short)MSVisio.VisRowIndices.visRowTextXForm,
                        (short)MSVisio.VisRowTags.visTagDefault);
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow(ex.ToString(), force: true);
                }
            }
        }

        public static void Validate_User_SectionExists(MSVisio.Shape shape)
        {
            if (0 == shape.SectionExists[(short)MSVisio.VisSectionIndices.visSectionUser, 0])
            {
                try
                {
                    var result = shape.AddSection((short)MSVisio.VisSectionIndices.visSectionUser);
                }
                catch (Exception ex)
                {
                    Common.WriteToDebugWindow(ex.ToString(), force: true);
                }
            }
        }
        private static void Set_RowXFormOut_Cell(
            MSVisio.Shape shape, 
            MSVisio.VisCellIndices cellIndex, 
            string value)
        {
            if (value != null)
            {
                shape.CellsSRC[
                    (short)MSVisio.VisSectionIndices.visSectionObject,
                    (short)MSVisio.VisRowIndices.visRowXFormOut,
                    (short)cellIndex].FormulaU = value;
            }
        }
    }
}
