using System;
using System.Windows;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class MiscellaneousRow
    {
        public string NoObjHandles { get; set; }
        public string NoCtlHandle { get; set; }
        public string NoAlignBox { get; set; }
        public string NonPrinting { get; set; }
        public string LangID { get; set; }
        public string HideText { get; set; }
        public string UpdateAlignBox { get; set; }
        public string DynFeedback { get; set; }
        public string NoLiveDynamics { get; set; }
        public string Calendar { get; set; }
        public string ObjType { get; set; }
        public string IsDropSource { get; set; }
        public string Comment { get; set; }
        public string DropOnPageScale { get; set; }
        public string LocalizeMerge { get; set; }
        public string NoProofing { get; set; }


        public static MiscellaneousRow GetRow(Shape shape)
        {
            MiscellaneousRow row = new MiscellaneousRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

            if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowMisc))
            {
                MessageBox.Show("No visRowMisc exists");
            }
            else
            {
                Row sectionRow = section[(short)VisRowIndices.visRowMisc];

                row.NoObjHandles = sectionRow[VisCellIndices.visNoObjHandles].FormulaU;
                row.NoCtlHandle = sectionRow[VisCellIndices.visNoCtlHandles].FormulaU;
                row.NoAlignBox = sectionRow[VisCellIndices.visNoAlignBox].FormulaU;
                row.NonPrinting = sectionRow[VisCellIndices.visNonPrinting].FormulaU;
                row.LangID = sectionRow[VisCellIndices.visObjLangID].FormulaU;
                row.HideText = sectionRow[VisCellIndices.visHideText].FormulaU;
                row.UpdateAlignBox = sectionRow[VisCellIndices.visUpdateAlignBox].FormulaU;
                row.DynFeedback = sectionRow[VisCellIndices.visDynFeedback].FormulaU;
                row.NoLiveDynamics = sectionRow[VisCellIndices.visNoLiveDynamics].FormulaU;
                row.Calendar = sectionRow[VisCellIndices.visObjCalendar].FormulaU;
                row.ObjType = sectionRow[VisCellIndices.visLOFlags].FormulaU;
                row.IsDropSource = sectionRow[VisCellIndices.visDropSource].FormulaU;
                row.Comment = sectionRow[VisCellIndices.visComment].FormulaU;
                row.DropOnPageScale = sectionRow[VisCellIndices.visObjDropOnPageScale].FormulaU;
                row.LocalizeMerge = sectionRow[VisCellIndices.visObjLocalizeMerge].FormulaU;
                row.NoProofing = sectionRow[VisCellIndices.visObjNoProofing].FormulaU;
            }
            return row;
        }

        public static void SetRow(Shape shape, MiscellaneousRow miscellaneous)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];

                if (Helpers.RowExists(shape, VisSectionIndices.visSectionObject, VisRowIndices.visRowMisc))
                {
                    MessageBox.Show("No visRowMisc exists");
                }
                else
                {
                    Row sectionRow = section[(short)VisRowIndices.visRowMisc];

                    sectionRow[VisCellIndices.visNoObjHandles].FormulaU = miscellaneous.NoObjHandles;
                    sectionRow[VisCellIndices.visNoCtlHandles].FormulaU = miscellaneous.NoCtlHandle;
                    sectionRow[VisCellIndices.visNoAlignBox].FormulaU = miscellaneous.NoAlignBox;
                    sectionRow[VisCellIndices.visNonPrinting].FormulaU = miscellaneous.NonPrinting;
                    sectionRow[VisCellIndices.visObjLangID].FormulaU = miscellaneous.LangID;
                    sectionRow[VisCellIndices.visHideText].FormulaU = miscellaneous.HideText;
                    sectionRow[VisCellIndices.visUpdateAlignBox].FormulaU = miscellaneous.UpdateAlignBox;
                    sectionRow[VisCellIndices.visDynFeedback].FormulaU = miscellaneous.DynFeedback;
                    sectionRow[VisCellIndices.visNoLiveDynamics].FormulaU = miscellaneous.NoLiveDynamics;
                    sectionRow[VisCellIndices.visObjCalendar].FormulaU = miscellaneous.Calendar;
                    sectionRow[VisCellIndices.visLOFlags].FormulaU = miscellaneous.ObjType;
                    sectionRow[VisCellIndices.visDropSource].FormulaU = miscellaneous.IsDropSource;
                    sectionRow[VisCellIndices.visComment].FormulaU = miscellaneous.Comment;
                    sectionRow[VisCellIndices.visObjDropOnPageScale].FormulaU = miscellaneous.DropOnPageScale;
                    sectionRow[VisCellIndices.visObjLocalizeMerge].FormulaU = miscellaneous.LocalizeMerge;
                    sectionRow[VisCellIndices.visObjNoProofing].FormulaU = miscellaneous.NoProofing;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
