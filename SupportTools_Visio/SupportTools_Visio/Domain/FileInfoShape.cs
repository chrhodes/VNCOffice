using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Visio = Microsoft.Office.Interop.Visio;

using Microsoft.CodeAnalysis;
using System.IO;
using Microsoft.CodeAnalysis.VisualBasic;
using VNC.CodeAnalysis;

namespace SupportTools_Visio.Domain
{
    public class FileInfoShape
    {
        delegate StringBuilder SearchTreeCommand(StringBuilder sb, SyntaxTree tree);


        #region Constructors and Load

        public FileInfoShape(Visio.Shape activeShape)
        {
            // TODO(crhodes)
            // Make this reflect on properties and loop across.

            BranchName = Helper.GetShapePropertyAsString(activeShape, "BranchName");
            BranchRepository = Helper.GetShapePropertyAsString(activeShape, "BranchRepository");
            BranchSourcePath = Helper.GetShapePropertyAsString(activeShape, "BranchSourcePath");
            SolutionName = Helper.GetShapePropertyAsString(activeShape, "SolutionName");
            SolutionFolderPath = Helper.GetShapePropertyAsString(activeShape, "SolutionFolderPath");
            SolutionFileName = Helper.GetShapePropertyAsString(activeShape, "SolutionFileName");
            ProjectName = Helper.GetShapePropertyAsString(activeShape, "ProjectName");
            ProjectFolderPath = Helper.GetShapePropertyAsString(activeShape, "ProjectFolderPath");
            ProjectFileName = Helper.GetShapePropertyAsString(activeShape, "ProjectFileName");
            SourceFileFolderPath = Helper.GetShapePropertyAsString(activeShape, "SourceFileFolderPath");
            SourceFileFileName = Helper.GetShapePropertyAsString(activeShape, "SourceFileFileName");
        }

        #endregion

        #region Enums, Fields, Properties, Structures

        public string BranchName { get; set; }
        public string BranchRepository { get; set; }
        public string BranchSourcePath { get; set; }
        public string SolutionName { get; set; }
        public string SolutionFolderPath { get; set; }
        public string SolutionFileName { get; set; }
        public string ProjectName { get; set; }
        public string ProjectFolderPath { get; set; }
        public string ProjectFileName { get; set; }
        public string SourceFileFolderPath { get; set; }
        public string SourceFileFileName { get; set; }

        #endregion

        #region Main Methods

        public string DisplayInfo()
        {
            StringBuilder sb = new StringBuilder();

            //ProcessOperation(sb, DisplayImportsStatementWalkerVB);
            //ProcessOperation(sb, DisplayNamespaceStatementWalkerVB);
            //ProcessOperation(sb, DisplayClassStatementWalkerVB);
            //ProcessOperation(sb, DisplayModuleStatementWalkerVB);
            //ProcessOperation(sb, DisplayMethodStatementWalkerVB);

            return sb.ToString();
        }

        //void CreateMethodShapes(Visio.Page page, Visio.Shape shape)
        //{
        //    ProcessOperation(sb, DisplayMethodStatementWalkerVB);
        //}

        public override string ToString()
        {
            return string.Format("{0} {1} {2} {3} {4} {5} {6} {7} {8} {9} {10}",
                BranchName, BranchRepository, BranchSourcePath, 
                SolutionName, SolutionFolderPath, SolutionFileName, 
                ProjectName, ProjectFolderPath, ProjectFileName,
                SourceFileFolderPath, SourceFileFileName);
        }

        #endregion

        private StringBuilder DisplayClassStatementWalkerVB(StringBuilder sb, SyntaxTree tree)
        {
            VNC.CodeAnalysis.SyntaxWalkers.VB.VNCVBSyntaxWalkerBase walker = null;

            //if ((bool)ceShowClassBlock.IsChecked)
            //{
            //    walker = new VNC.CodeAnalysis.SyntaxWalkers.VB.ClassBlock();
            //}
            //else
            //{
                walker = new VNC.CodeAnalysis.SyntaxWalkers.VB.ClassStatement();
            //}

            //return InvokeVNCSyntaxWalker(sb,
            //    (bool)ceClassStatementUseRegEx.IsChecked, teClassStatementRegEx.Text,
            //    tree, walker);

            return InvokeVNCSyntaxWalker(sb,
                false, "",
                tree, walker);
        }

        private StringBuilder DisplayImportsStatementWalkerVB(StringBuilder sb, SyntaxTree tree)
        {
            var walker = new VNC.CodeAnalysis.SyntaxWalkers.VB.ImportsStatement();

            //return InvokeVNCSyntaxWalker(sb,
            //    (bool)ceImportsStatementUseRegEx.IsChecked, teImportsStatementRegEx.Text,
            //    tree, walker);

            return InvokeVNCSyntaxWalker(sb,
                false, "",
                tree, walker);
        }

        private StringBuilder DisplayModuleStatementWalkerVB(StringBuilder sb, SyntaxTree tree)
        {
            VNC.CodeAnalysis.SyntaxWalkers.VB.VNCVBTypedSyntaxWalkerBase walker = null;

            //if ((bool)ceShowModuleBlock.IsChecked)
            //{
            //    walker = new VNC.CodeAnalysis.SyntaxWalkers.VB.ModuleBlock();
            //}
            //else
            //{
                walker = new VNC.CodeAnalysis.SyntaxWalkers.VB.ModuleStatement();
            //}

            //return InvokeVNCSyntaxWalker(sb,
            //    (bool)ceModuleStatementUseRegEx.IsChecked, teModuleStatementRegEx.Text,
            //    tree, walker);

            return InvokeVNCSyntaxWalker(sb,
                false, "",
                tree, walker);
        }

        private StringBuilder DisplayNamespaceStatementWalkerVB(StringBuilder sb, SyntaxTree tree)
        {
            var walker = new VNC.CodeAnalysis.SyntaxWalkers.VB.NamespaceStatement();

            //return InvokeVNCSyntaxWalker(sb,
            //    (bool)ceNamespaceStatementUseRegEx.IsChecked, teNamespaceStatementRegEx.Text,
            //    tree, walker);

            return InvokeVNCSyntaxWalker(sb,
                false, "",
                tree, walker);
        }

        private StringBuilder DisplayMethodStatementWalkerVB(StringBuilder sb, SyntaxTree tree)
        {
            VNC.CodeAnalysis.SyntaxWalkers.VB.VNCVBTypedSyntaxWalkerBase walker = null;

            //if ((bool)ceShowMethodBlock.IsChecked)
            //{
            //    walker = new VNC.CodeAnalysis.SyntaxWalkers.VB.MethodBlock();
            //}
            //else
            //{
                walker = new VNC.CodeAnalysis.SyntaxWalkers.VB.MethodStatement();
            //}

            //return InvokeVNCSyntaxWalker(sb,
            //    (bool)ceMethodStatementUseRegEx.IsChecked, teMethodStatementRegEx.Text,
            //    tree, walker);

            return InvokeVNCSyntaxWalker(sb,
               false, "",
                tree, walker);
        }

        private StringBuilder ProcessOperation()
        {
            StringBuilder sb = new StringBuilder();
            //CodeExplorer.teSourceCode.Clear();

            //string projectFullPath = CodeExplorerContext.teProjectFile.Text;
            //string pattern = teInvocationExpressionRegEx.Text;

            //var filesToProcess = CodeExplorerContext.GetFilesToProcess();

            //if (filesToProcess.Count > 0)
            //{
            //    if ((Boolean)ceListImpactedFilesOnly.IsChecked)
            //    {
            //        sb.AppendLine("Would Search these files ....");
            //    }

            //    foreach (string filePath in filesToProcess)
            //    {
            //        if ((Boolean)ceListImpactedFilesOnly.IsChecked)
            //        {
            //            sb.AppendLine(string.Format("  {0}", filePath));
            //        }
            //        else
            //        {
            //StringBuilder sbFileResults = new StringBuilder();

            //var sourceCode = "";

            //using (var sr = new StreamReader(SourceFileFileName))
            //{
            //    sourceCode = sr.ReadToEnd();
            //}

            //SyntaxTree tree = VisualBasicSyntaxTree.ParseText(sourceCode);

            //sbFileResults = command(sbFileResults, tree);

            ////if ((bool)ceAlwaysDisplayFileName.IsChecked || (sbFileResults.Length > 0))
            ////{
            ////    sb.AppendLine("Searching " + filePath);
            ////}

            //sb.Append(sbFileResults.ToString());
            //        }
            //    }
            //}
            //else
            //{
            //    sb.AppendLine("No files selected to process");
            //}

            return sb;
            //CodeExplorer.teSourceCode.Text = sb.ToString();
        }

        StringBuilder InvokeVNCSyntaxWalker(
            StringBuilder sb,
            Boolean useRegEx, string regEx,
            SyntaxTree syntaxTree,
            VNC.CodeAnalysis.SyntaxWalkers.VB.VNCVBSyntaxWalkerBase walker)
        {
            walker.Messages = sb;

            //walker.DisplayClassOrModuleName = (bool)ceDisplayClassOrModuleName.IsChecked;
            //walker.DisplayMethodName = (bool)ceDisplayMethodName.IsChecked;
            
            //walker.DisplayClassOrModuleName = true;
            //walker.DisplayMethodName = true;

            //if (useRegEx)
            //{
            //    walker.IdentifierNames = regEx;
            //}
            //else
            //{
            //    walker.IdentifierNames = ".*";
            //}

            //walker.InitializeRegEx();

            //walker.Visit(syntaxTree.GetRoot());

            return sb;
        }
    }
}
