using System;

using Microsoft.Office.Interop.Visio;

namespace VNC.Visio.VSTOAddIn.Domain
{
    public class ProtectionRow
    {
        public string LockWidth { get; set; }
        public string LockHeight { get; set; }
        public string LockAspect { get; set; }
        public string LockMoveX { get; set; }
        public string LockMoveY { get; set; }
        public string LockRotate { get; set; }
        public string LockBegin { get; set; }
        public string LockReplace { get; set; }
        public string LockEnd { get; set; }
        public string LockDelete { get; set; }
        public string LockSelect { get; set; }
        public string LockFormat { get; set; }
        public string LockCustProp { get; set; }
        public string LockTextEdit { get; set; }
        public string LockVtxEdit { get; set; }
        public string LockThemeIndex { get; set; }
        public string LockCrop { get; set; }
        public string LockGroup { get; set; }
        public string LockCalcWH { get; set; }
        public string LockFromGroupFormat { get; set; }
        public string LockThemeColors { get; set; }
        public string LockThemeEffects { get; set; }
        public string LockThemeConnectors { get; set; }
        public string LockThemeFonts { get; set; }
        public string LockVariation { get; set; }

        public static ProtectionRow Get_Protection(Shape shape)
        {
            ProtectionRow row = new ProtectionRow();

            Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
            Row sectionRow = section[(short)VisRowIndices.visRowLock];

            row.LockWidth = sectionRow[VisCellIndices.visLockWidth].FormulaU;
            row.LockHeight = sectionRow[VisCellIndices.visLockHeight].FormulaU;
            row.LockAspect = sectionRow[VisCellIndices.visLockAspect].FormulaU;
            row.LockMoveX = sectionRow[VisCellIndices.visLockMoveX].FormulaU;
            row.LockMoveY = sectionRow[VisCellIndices.visLockMoveY].FormulaU;
            row.LockRotate = sectionRow[VisCellIndices.visLockRotate].FormulaU;
            row.LockBegin = sectionRow[VisCellIndices.visLockBegin].FormulaU;
            row.LockReplace = sectionRow[VisCellIndices.visLockReplace].FormulaU;
            row.LockEnd = sectionRow[VisCellIndices.visLockEnd].FormulaU;
            row.LockDelete = sectionRow[VisCellIndices.visLockDelete].FormulaU;
            row.LockSelect = sectionRow[VisCellIndices.visLockSelect].FormulaU;
            row.LockFormat = sectionRow[VisCellIndices.visLockFormat].FormulaU;
            row.LockCustProp = sectionRow[VisCellIndices.visLockCustProp].FormulaU;
            row.LockTextEdit = sectionRow[VisCellIndices.visLockTextEdit].FormulaU;
            row.LockVtxEdit = sectionRow[VisCellIndices.visLockVtxEdit].FormulaU;
            row.LockThemeIndex = sectionRow[VisCellIndices.visLockThemeIndex].FormulaU;
            row.LockCrop = sectionRow[VisCellIndices.visLockCrop].FormulaU;
            row.LockGroup = sectionRow[VisCellIndices.visLockGroup].FormulaU;
            row.LockCalcWH = sectionRow[VisCellIndices.visLockCalcWH].FormulaU;
            row.LockFromGroupFormat = sectionRow[VisCellIndices.visLockFromGroupFormat].FormulaU;
            row.LockThemeColors = sectionRow[VisCellIndices.visLockThemeColors].FormulaU;
            row.LockThemeEffects = sectionRow[VisCellIndices.visLockThemeEffects].FormulaU;
            row.LockThemeConnectors = sectionRow[VisCellIndices.visLockThemeConnectors].FormulaU;
            row.LockThemeFonts = sectionRow[VisCellIndices.visLockThemeFonts].FormulaU;
            row.LockVariation = sectionRow[VisCellIndices.visLockVariation].FormulaU;

            return row;
        }

        public static void Set_Protection_Section(Shape shape, ProtectionRow protection)
        {
            try
            {
                Section section = shape.Section[(short)VisSectionIndices.visSectionObject];
                Row sectionRow = section[(short)VisRowIndices.visRowLock];

                sectionRow[VisCellIndices.visLockWidth].FormulaU = protection.LockWidth;
                sectionRow[VisCellIndices.visLockHeight].FormulaU = protection.LockHeight;
                sectionRow[VisCellIndices.visLockAspect].FormulaU = protection.LockAspect;
                sectionRow[VisCellIndices.visLockMoveX].FormulaU = protection.LockMoveX;
                sectionRow[VisCellIndices.visLockMoveY].FormulaU = protection.LockMoveY;
                sectionRow[VisCellIndices.visLockRotate].FormulaU = protection.LockRotate;
                sectionRow[VisCellIndices.visLockBegin].FormulaU = protection.LockBegin;
                sectionRow[VisCellIndices.visLockReplace].FormulaU = protection.LockReplace;
                sectionRow[VisCellIndices.visLockEnd].FormulaU = protection.LockEnd;
                sectionRow[VisCellIndices.visLockDelete].FormulaU = protection.LockDelete;
                sectionRow[VisCellIndices.visLockSelect].FormulaU = protection.LockSelect;
                sectionRow[VisCellIndices.visLockFormat].FormulaU = protection.LockFormat;
                sectionRow[VisCellIndices.visLockCustProp].FormulaU = protection.LockCustProp;
                sectionRow[VisCellIndices.visLockTextEdit].FormulaU = protection.LockTextEdit;
                sectionRow[VisCellIndices.visLockVtxEdit].FormulaU = protection.LockVtxEdit;
                sectionRow[VisCellIndices.visLockThemeIndex].FormulaU = protection.LockThemeIndex;
                sectionRow[VisCellIndices.visLockCrop].FormulaU = protection.LockCrop;
                sectionRow[VisCellIndices.visLockGroup].FormulaU = protection.LockGroup;
                sectionRow[VisCellIndices.visLockCalcWH].FormulaU = protection.LockCalcWH;
                sectionRow[VisCellIndices.visLockFromGroupFormat].FormulaU = protection.LockFromGroupFormat;
                sectionRow[VisCellIndices.visLockThemeColors].FormulaU = protection.LockThemeColors;
                sectionRow[VisCellIndices.visLockThemeEffects].FormulaU = protection.LockThemeEffects;
                sectionRow[VisCellIndices.visLockThemeConnectors].FormulaU = protection.LockThemeConnectors;
                sectionRow[VisCellIndices.visLockThemeFonts].FormulaU = protection.LockThemeFonts;
                sectionRow[VisCellIndices.visLockVariation].FormulaU = protection.LockVariation;
            }
            catch (Exception ex)
            {
                Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
