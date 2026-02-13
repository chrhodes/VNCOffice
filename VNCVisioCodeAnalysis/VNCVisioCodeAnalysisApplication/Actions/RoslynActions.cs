using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

using VNCVisioToolsApplication.Domain;

using VNC.Core;

using MSVisio = Microsoft.Office.Interop.Visio;
using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace VNCVisioToolsApplication.Actions
{
    public class RoslynActions
    {

        public static void GetClassInfo(MSVisio.Application app, string doc, string page, string shape, string shapeu, String[] args)
        {
            Common.WriteToDebugWindow(string.Format("{0}()",
                MethodBase.GetCurrentMethod().Name));

            MSVisio.Page activePage = app.ActivePage;
            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];

            ClassInfoShape classInfoShape = new ClassInfoShape(activeShape);

            VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}",
                classInfoShape.ToString()));
        }

        internal static void GetProjectFileInfo(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] v)
        {
            VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}()",
                MethodBase.GetCurrentMethod().Name));

            MSVisio.Page activePage = app.ActivePage;
            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];
            List<String> filesToProcess; // = new List<string>();

            try
            {
                FileInfoShape fileInfoShape = new FileInfoShape(activeShape);

                VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}",
                    fileInfoShape.ToString()));

                //VNCVisioAddIn.Common.DisplayInDebugWindow(string.Format("{0}",
                //    fileInfoShape.DisplayInfo()));

                string configFileFullPath = @"C:\temp\CodeCommandConsole_Config.xml";

                string projectFileName = fileInfoShape.ProjectFileName;
                string solutionFileName = fileInfoShape.SolutionFileName;
                string branchName = fileInfoShape.BranchName;
                string solutionName = fileInfoShape.SolutionName;
                string projectName = fileInfoShape.ProjectName;


                if (projectFileName.Length > 0)
                {
                    string sourcePath = fileInfoShape.BranchSourcePath;
                    string projectFolderPath = fileInfoShape.ProjectFolderPath;

                    filesToProcess = VNC.CodeAnalysis.Workspace.Helper.GetSourceFilesToProcessFromVSProject(
                        sourcePath + "\\" + projectFolderPath + "\\" + projectFileName);
                }
                else
                {
                    filesToProcess = VNC.CodeAnalysis.Workspace.Helper.GetSourceFilesToProcessFromConfigFile(configFileFullPath, branchName, solutionName, projectName);
                }

                foreach (string filePath in filesToProcess)
                {
                    VNCVisioAddIn.Common.WriteToDebugWindow($"{filePath}");
                    //if ((Boolean)ceListImpactedFilesOnly.IsChecked)
                    //{
                    //    sb.AppendLine(string.Format("  {0}", filePath));
                    //}
                    //else
                    //{
                    //    StringBuilder sbFileResults = new StringBuilder();

                    //    var sourceCode = "";

                    //    using (var sr = new StreamReader(filePath))
                    //    {
                    //        sourceCode = sr.ReadToEnd();
                    //    }

                    //    SyntaxTree tree = VisualBasicSyntaxTree.ParseText(sourceCode);

                    //    sbFileResults = command(sbFileResults, matches, tree);

                    //    if ((bool)ceAlwaysDisplayFileName.IsChecked || (sbFileResults.Length > 0))
                    //    {
                    //        sb.AppendLine("Searching " + filePath);
                    //    }

                    //    sb.Append(sbFileResults.ToString());
                    //}
                }

                //var sourceCode = "";

                //using (var sr = new StreamReader(fileNameAndPath))
                //{
                //    sourceCode = sr.ReadToEnd();
                //}

                //List<String> methodNames = VNC.CodeAnalysis.Helpers.VB.GetMethodNames(sourceCode);

                //// OK.  Now we have a list of Method Names.  Let's add shapes for each of them.

                //Visio.Master methodMaster = app.Documents[@"API.vssx"].Masters[@"Roslyn SuperFile"];

                //foreach (string methodName in methodNames)
                //{
                //    Visio.Shape newMethod = activePage.Drop(methodMaster, 5.0, 5.0);
                //    newMethod.CellsU["Prop.MethodName"].FormulaU = methodName.WrapInDblQuotes();
                //}
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}",
                    ex.ToString()));
            }
        }

        internal static void GetSolutionFileInfo(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] v)
        {
            VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}()",
                MethodBase.GetCurrentMethod().Name));

            MSVisio.Page activePage = app.ActivePage;
            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];

            try
            {
                FileInfoShape fileInfoShape = new FileInfoShape(activeShape);

                VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}",
                    fileInfoShape.ToString()));

                VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}",
                    fileInfoShape.DisplayInfo()));

                string fileNameAndPath = fileInfoShape.SolutionFileName;

                var sourceCode = "";

                using (var sr = new StreamReader(fileNameAndPath))
                {
                    sourceCode = sr.ReadToEnd();
                }

                List<String> methodNames = VNC.CodeAnalysis.Helpers.VB.GetMethodNames(sourceCode);

                // OK.  Now we have a list of Method Names.  Let's add shapes for each of them.

                MSVisio.Master methodMaster = app.Documents[@"API.vssx"].Masters[@"Roslyn SuperFile"];

                foreach (string methodName in methodNames)
                {
                    MSVisio.Shape newMethod = activePage.Drop(methodMaster, 5.0, 5.0);
                    newMethod.CellsU["Prop.MethodName"].FormulaU = methodName.WrapInDblQuotes();
                }
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}",
                    ex.ToString()));
            }
        }

        internal static void GetSourceFileInfo(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] v)
        {
            VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}()",
                MethodBase.GetCurrentMethod().Name));

            MSVisio.Page activePage = app.ActivePage;
            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];

            try
            {
                FileInfoShape fileInfoShape = new FileInfoShape(activeShape);

                VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}",
                    fileInfoShape.ToString()));

                VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}",
                    fileInfoShape.DisplayInfo()));

                string fileNameAndPath = fileInfoShape.SourceFileFileName;

                var sourceCode = "";

                using (var sr = new StreamReader(fileNameAndPath))
                {
                    sourceCode = sr.ReadToEnd();
                }

                List<String> methodNames = VNC.CodeAnalysis.Helpers.VB.GetMethodNames(sourceCode);

                // OK.  Now we have a list of Method Names.  Let's add shapes for each of them.

                MSVisio.Master methodMaster = app.Documents[@"API.vssx"].Masters[@"Roslyn SuperFile"];

                foreach (string methodName in methodNames)
                {
                    MSVisio.Shape newMethod = activePage.Drop(methodMaster, 5.0, 5.0);
                    newMethod.CellsU["Prop.MethodName"].FormulaU = methodName.WrapInDblQuotes();
                }
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}",
                    ex.ToString()));
            }
        }

        internal static void GetMethodInfo(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] v)
        {
            VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}()",
                MethodBase.GetCurrentMethod().Name));

            MSVisio.Page activePage = app.ActivePage;
            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];

            MethodInfoShape methodInfoShape = new MethodInfoShape(activeShape);

            VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}",
                methodInfoShape.ToString()));
        }

        internal static void CreateMethodShapes(MSVisio.Application app, string doc, string page, string shape, string shapeu, string[] v)
        {
            VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}()",
                MethodBase.GetCurrentMethod().Name));

            MSVisio.Page activePage = app.ActivePage;
            MSVisio.Shape activeShape = app.ActivePage.Shapes[shape];

            try
            {
                FileInfoShape fileInfoShape = new FileInfoShape(activeShape);


                VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}",
                    fileInfoShape.ToString()));

                VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}",
                    fileInfoShape.DisplayInfo()));

                string fileNameAndPath = fileInfoShape.SourceFileFileName;

                var sourceCode = "";

                using (var sr = new StreamReader(fileNameAndPath))
                {
                    sourceCode = sr.ReadToEnd();
                }

                List<String> methodNames = VNC.CodeAnalysis.Helpers.VB.GetMethodNames(sourceCode);

                // OK.  Now we have a list of Method Names.  Let's add shapes for each of them.

                MSVisio.Master methodMaster = app.Documents[@"API.vssx"].Masters[@"Method"];

                foreach (string methodName in methodNames)
                {
                    MSVisio.Shape newMethod = activePage.Drop(methodMaster, 5.0, 5.0);
                    newMethod.CellsU["Prop.MethodName"].FormulaU = methodName.WrapInDblQuotes();
                }
            }
            catch (Exception ex)
            {
                VNCVisioAddIn.Common.WriteToDebugWindow(string.Format("{0}",
                    ex.ToString()));
            }
        }
    }
}
