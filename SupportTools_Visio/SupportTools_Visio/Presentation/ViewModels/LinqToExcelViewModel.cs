using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;

using Prism.Commands;
using Prism.Events;
using Prism.Services.Dialogs;

using SupportTools_Visio.Core;

using VNC;
using VNC.Core.Mvvm;

using ExcelHlp = VNC.AddinHelper.Excel;
using LTE = LinqToExcel;
using VisioHlp = VNC.AddinHelper.Visio;

using XL = Microsoft.Office.Interop.Excel;

namespace SupportTools_Visio.Presentation.ViewModels
{
    public class TestData
    {
        public string Col1 { get; set; }
        public string Col2 { get; set; }
        public string Col3 { get; set; }
        public string Col4 { get; set; }
        public string Col5 { get; set; }
    }
    public class LinqToExcelViewModel : EventViewModelBase, ILinqToExcelViewModel, IInstanceCountVM
    {

        #region Constructors, Initialization, and Load

        public LinqToExcelViewModel(
            IEventAggregator eventAggregator,
            DialogService dialogService) : base(eventAggregator, dialogService)
        {
            Int64 startTicks = Log.CONSTRUCTOR("Enter", Common.LOG_CATEGORY);

            // TODO(crhodes)
            // Save constructor parameters here

            InitializeViewModel();

            Log.CONSTRUCTOR("Exit", Common.LOG_CATEGORY, startTicks);
        }

        private void InitializeViewModel()
        {
            Int64 startTicks = Log.VIEWMODEL("Enter", Common.LOG_CATEGORY);

            InstanceCountVM++;

            // TODO(crhodes)
            //

            SayHelloCommand = new DelegateCommand(
                SayHello, SayHelloCanExecute);

            UseLinqToExcelCommand = new DelegateCommand(
                UseLinqToExcel, UseLinqToExcelCanExecute);

            LoadExcelTableCommand = new DelegateCommand(
                LoadExcelTable, LoadExcelTableCanExecute);

            UseExcelDataReaderCommand = new DelegateCommand(
                UseExcelDataReader, UseExcelDataReaderCanExecute);

            LoadExcelFileCommand = new DelegateCommand(
                LoadExcelFile, LoadExcelFileCanExecute);

            Message = "LinqToExcelViewModel says hello";

            Log.VIEWMODEL("Exit", Common.LOG_CATEGORY, startTicks);
        }

        #endregion

        #region Enums


        #endregion

        #region Structures


        #endregion

        #region Fields and Properties

        public ICommand SayHelloCommand { get; private set; }

        private string _message;

        public string Message
        {
            get => _message;
            set
            {
                if (_message == value)
                    return;
                _message = value;
                OnPropertyChanged();
            }
        }

        #endregion

        #region Event Handlers



        #endregion

        #region Public Methods


        #endregion

        #region Protected Methods


        #endregion

        #region Private Methods

        #region UseLinqToExcel Command

        public DelegateCommand UseLinqToExcelCommand { get; set; }
        public string UseLinqToExcelContent { get; set; } = "UseLinqToExcel";
        public string UseLinqToExcelToolTip { get; set; } = "UseLinqToExcel ToolTip";

        // Can get fancy and use Resources
        //public string UseLinqToExcelContent { get; set; } = "ViewName_UseLinqToExcelContent";
        //public string UseLinqToExcelToolTip { get; set; } = "ViewName_UseLinqToExcelContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_UseLinqToExcelContent">UseLinqToExcel</system:String>
        //    <system:String x:Key="ViewName_UseLinqToExcelContentToolTip">UseLinqToExcel ToolTip</system:String>  

        public void UseLinqToExcel()
        {
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called UseLinqToExcel";
            Common.EventAggregator.GetEvent<UseLinqToExcelEvent>().Publish();

        // Put this in places that listen for event
        //Common.EventAggregator.GetEvent<UseLinqToExcelEvent>().Subscribe(UseLinqToExcel);
        }

        private void UseLinqToExcelX()
        {
            string path = @"B:\Publish\SupportTools_Visio\TestData.xlsx";
            var excel = new LTE.ExcelQueryFactory(path);

            var stuff = from c in excel.Worksheet<TestData>()
                        select c;

            foreach (var item in stuff)
            {
                VisioHlp.DisplayInWatchWindow(
                    string.Format("Col1:{0} Col2:{1} Col3:{2} Col4:{3} Col5:{4}", item.Col1, item.Col2, item.Col3, item.Col4, item.Col5)
                    );
            }
        }

        public bool UseLinqToExcelCanExecute()
    {
        // TODO(crhodes)
        // Add any before button is enabled logic.
        return true;
    }

        #endregion

        #region LoadExcelTable Command

        public DelegateCommand LoadExcelTableCommand { get; set; }
        public string LoadExcelTableContent { get; set; } = "LoadExcelTable";
        public string LoadExcelTableToolTip { get; set; } = "LoadExcelTable ToolTip";

        // Can get fancy and use Resources
        //public string LoadExcelTableContent { get; set; } = "ViewName_LoadExcelTableContent";
        //public string LoadExcelTableToolTip { get; set; } = "ViewName_LoadExcelTableContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_LoadExcelTableContent">LoadExcelTable</system:String>
        //    <system:String x:Key="ViewName_LoadExcelTableContentToolTip">LoadExcelTable ToolTip</system:String>  

        public void LoadExcelTable()
        {
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called LoadExcelTable";
            Common.EventAggregator.GetEvent<LoadExcelTableEvent>().Publish();

            // Start Cut Four

            // Put this in places that listen for event
            //Common.EventAggregator.GetEvent<LoadExcelTableEvent>().Subscribe(LoadExcelTable);

            // End Cut Four
        }

    private void LoadExcelTableX()
    {
        string workBookName = @"B:\Publish\SupportTools_Visio\TestData.xlsx";
        string workSheetName = "Sheet2";
        string tableName = "tbl_Data";
        XL.Application xlApp = new XL.Application();

        XL.Workbook wb = xlApp.Workbooks.Open(workBookName);
        XL.Worksheet ws = wb.Sheets[workSheetName];
        XL.ListObject lo = ws.ListObjects[tableName];
        XL.ListColumns listColumns = lo.ListColumns;
        XL.ListRows listRows = lo.ListRows;

        VisioHlp.DisplayInWatchWindow(string.Format("{0}\n", tableName));

        foreach (XL.ListColumn col in listColumns)
        {
            VisioHlp.DisplayInWatchWindow(col.Name);
        }

        tableName = "tbl_Data2";

        lo = ws.ListObjects[tableName];
        listColumns = lo.ListColumns;
        listRows = lo.ListRows;

        VisioHlp.DisplayInWatchWindow(string.Format("{0}\n", tableName));

        foreach (XL.ListColumn col in listColumns)
        {
            VisioHlp.DisplayInWatchWindow(col.Name);
        }

        foreach (XL.ListRow row in listRows)
        {
            VisioHlp.DisplayInWatchWindow(row.ToString());
        }
        wb.Close();
    }

    public bool LoadExcelTableCanExecute()
    {
        // TODO(crhodes)
        // Add any before button is enabled logic.
        return true;
    }

        #endregion

        #region UseExcelDataReader Command

        public DelegateCommand UseExcelDataReaderCommand { get; set; }
        public string UseExcelDataReaderContent { get; set; } = "UseExcelDataReader";
        public string UseExcelDataReaderToolTip { get; set; } = "UseExcelDataReader ToolTip";

        // Can get fancy and use Resources
        //public string UseExcelDataReaderContent { get; set; } = "ViewName_UseExcelDataReaderContent";
        //public string UseExcelDataReaderToolTip { get; set; } = "ViewName_UseExcelDataReaderContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_UseExcelDataReaderContent">UseExcelDataReader</system:String>
        //    <system:String x:Key="ViewName_UseExcelDataReaderContentToolTip">UseExcelDataReader ToolTip</system:String>  

        public void UseExcelDataReader()
        {
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called UseExcelDataReader";
            Common.EventAggregator.GetEvent<UseExcelDataReaderEvent>().Publish();

            // Start Cut Four

            // Put this in places that listen for event
            //Common.EventAggregator.GetEvent<UseExcelDataReaderEvent>().Subscribe(UseExcelDataReader);

            // End Cut Four

        }

        private void UseExcelDataReaderX()
        {
            string path = @"B:\Publish\SupportTools_Visio\TestData.xlsx";

            var excelData = new ExcelHlp.XlData(path);
            var sheets = excelData.GetWorkSheetNames();

            foreach (var sheet in sheets)
            {
                VisioHlp.DisplayInWatchWindow(sheet);
            }

            VisioHlp.DisplayInWatchWindow("Has Header Row\n");

            var info = excelData.GetData("Sheet1");

            foreach (var row in info)
            {
                VisioHlp.DisplayInWatchWindow("NewRow\n");

                for (int i = 0; i <= row.ItemArray.GetUpperBound(0); i++)
                {
                    VisioHlp.DisplayInWatchWindow(row[i].ToString());
                }
            }

            VisioHlp.DisplayInWatchWindow("Has No Header Row\n");

            info = excelData.GetData("Sheet1", false);

            foreach (var row in info)
            {
                VisioHlp.DisplayInWatchWindow("NewRow\n");

                for (int i = 0; i <= row.ItemArray.GetUpperBound(0); i++)
                {
                    VisioHlp.DisplayInWatchWindow(row[i].ToString());
                }
            }

            List<TestData> testData = new List<TestData>();

            info = excelData.GetData("Sheet1");

            foreach (var row in info)
            {
                var testDataRow = new TestData()
                {
                    Col1 = row["Col1"].ToString(),
                    Col2 = row["Col2"].ToString(),
                    Col3 = row["Col3"].ToString(),
                    Col4 = row["Col4"].ToString(),
                    Col5 = row["Col5"].ToString()
                };

                testData.Add(testDataRow);
            }

            foreach (var item in testData)
            {
                VisioHlp.DisplayInWatchWindow(
                    string.Format("Col1:{0} Col2:{1} Col3:{2} Col4:{3} Col5:{4}", item.Col1, item.Col2, item.Col3, item.Col4, item.Col5)
                    );
            }
        }


    public bool UseExcelDataReaderCanExecute()
    {
        // TODO(crhodes)
        // Add any before button is enabled logic.
        return true;
    }

        #endregion


        #region LoadExcelFile Command

        public DelegateCommand LoadExcelFileCommand { get; set; }
        public string LoadExcelFileContent { get; set; } = "LoadExcelFile";
        public string LoadExcelFileToolTip { get; set; } = "LoadExcelFile ToolTip";

        // Can get fancy and use Resources
        //public string LoadExcelFileContent { get; set; } = "ViewName_LoadExcelFileContent";
        //public string LoadExcelFileToolTip { get; set; } = "ViewName_LoadExcelFileContentToolTip";

        // Put these in Resource File
        //    <system:String x:Key="ViewName_LoadExcelFileContent">LoadExcelFile</system:String>
        //    <system:String x:Key="ViewName_LoadExcelFileContentToolTip">LoadExcelFile ToolTip</system:String>  

        public void LoadExcelFile()
        {
            // TODO(crhodes)
            // Do something amazing.
            Message = "Cool, you called LoadExcelFile";
            Common.EventAggregator.GetEvent<LoadExcelFileEvent>().Publish();

            // Start Cut Four

            // Put this in places that listen for event
            //Common.EventAggregator.GetEvent<LoadExcelFileEvent>().Subscribe(LoadExcelFile);

            // End Cut Four

        }

        private void LoadExcelFileX()
        {
            XL.Application xlApp = new XL.Application();

            XL.Workbook wb = xlApp.Workbooks.Open(@"B:\Publish\SupportTools_Visio\TestData.xlsx");
            XL.Worksheet ws = wb.Sheets[1];
            XL.Range rng = ws.UsedRange;

            int rows = rng.Rows.Count;
            int cols = rng.Columns.Count;

            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= cols; j++)
                {
                    VisioHlp.DisplayInWatchWindow(rng.Cells[i, j].Value2.ToString());
                }
            }

            wb.Close();
        }

        public bool LoadExcelFileCanExecute()
        {
            // TODO(crhodes)
            // Add any before button is enabled logic.
            return true;
        }

        #endregion

        private bool SayHelloCanExecute()
        {
            return true;
        }

        private void SayHello()
        {
            Int64 startTicks = Log.EVENT_HANDLER("Enter", Common.LOG_CATEGORY);

            Message = "Hello";

            Log.EVENT_HANDLER("Exit", Common.LOG_CATEGORY, startTicks);
        }



        #endregion

        #region IInstanceCount

        private static int _instanceCountVM;

        public int InstanceCountVM
        {
            get => _instanceCountVM;
            set => _instanceCountVM = value;
        }

        #endregion
    }
}
