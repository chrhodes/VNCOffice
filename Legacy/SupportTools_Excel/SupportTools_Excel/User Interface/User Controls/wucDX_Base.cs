using System;
using System.Windows;
using System.Windows.Controls;
//using DevExpress.Xpf.DemoBase;
//using EyeOnLife.User_Interface.Windows;
using DevExpress.Xpf.Grid;

namespace SupportTools_Excel.User_Interface.User_Controls
{
    public class wucDX_Base : UserControl, IRefresh
    {
        private const string LOG_APPNAME = "SupportTools_Excel";

        // These stand-in for the real controls that are added
        // on the pages that derive from wucBase
        // and allow the common code to be in this base class.

        protected GridControl dataGrid = null;
        protected GridColumn currentColumn = null;

        //protected TableView tableView = null;

        //protected GridColumn gc_ID = null;
        //protected GridColumn gc_SnapShotDate = null;
        //protected GridColumn gc_SnapShotDuration = null;
        //protected GridColumn gc_SnapShotError = null;

        public wucDX_Base()
        {
            Loaded += EditForm_Loaded;
        }

        internal bool _editMode = false;

        protected virtual void OnLoaded(object sender, RoutedEventArgs e)
        {
            
        }

        protected void EditForm_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            ApplyAuthorization();
        }

        void IRefresh.Refresh()
        {
            ApplyAuthorization();
        }

        /// <summary>
        /// Override this method and use to apply
        /// authorization rules as the form loads
        /// or the current user changes.
        /// </summary>
        protected virtual void ApplyAuthorization()
        {

        }

        #region Title property

        public event EventHandler TitleChanged;
        //public event EventHandler TitleChanged2;

        /// <summary>
        /// Gets or sets the title of this
        /// edit form.
        /// </summary>
        public string Title
        {
              get { return (string)GetValue(TitleProperty); }

              set
              {
                    //SetValue(TitleProperty, value);
                    //SetValue(TitleProperty, value);

                    if (TitleChanged != null)
                    {
                        TitleChanged(this, EventArgs.Empty);
                    }
              }
        }

        // Using a DependencyProperty as the backing store for _title.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty TitleProperty =
            DependencyProperty.Register("Title", typeof(string), typeof(wucDX_Base), null);

        #endregion

        protected void LogUsage(Type type)
        {
            var eventMessage = string.Format("OnLoaded: {0}", type.Name);
            //SQLInformation.Helper.IndicateApplicationUsage(LOG_APPNAME, DateTime.Now, Common.CurrentUser.Identity.Name, eventMessage);
        }

        protected virtual void DataChanged(object sender, EventArgs e)
        {
          var dp = sender as System.Windows.Data.DataSourceProvider;

          if (dp.Error != null)
          {
              MessageBox.Show(dp.Error.ToString(), "Data error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
          }
        }

        protected void canAddCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (dataGrid != null)
            {
                //dataGrid.CanUserAddRows = true;
            }
        }

        protected void canAddCheckBox_UnChecked(object sender, RoutedEventArgs e)
        {
            
            if (dataGrid != null)
            {
                //dataGrid.CanUserAddRows = false;
            }
        }

        protected void canDeleteCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (dataGrid != null)
            {
                //dataGrid.CanUserDeleteRows = true;
            }
        }

        protected void canDeleteCheckBox_UnChecked(object sender, RoutedEventArgs e)
        {
            if (dataGrid != null)
            {
                //dataGrid.CanUserDeleteRows = false;
            }
        }

        protected void readOnlyCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (dataGrid != null)
            {
                //dataGrid.IsReadOnly = true;
            }
        }

        protected void readOnlyCheckBox_UnChecked(object sender, RoutedEventArgs e)
        {
            if (dataGrid != null)
            {
                //dataGrid.IsReadOnly = false;
            }
        }

        //protected void OnDisplaySnapShotColumns_Checked(object sender, RoutedEventArgs e)
        //{
        //    gc_SnapShotDate.Visible = true;
        //    gc_SnapShotError.Visible = true;
        //}

        //protected void ckDisplaySnapShotColumns_Unchecked(object sender, RoutedEventArgs e)
        //{
        //    gc_SnapShotDate.Visible = false;
        //    gc_SnapShotError.Visible = false;
        //}

        //protected void OnDisplayIDColumns_Checked(object sender, RoutedEventArgs e)
        //{
        //    //HideIDColumns(((CheckBox)sender).IsChecked);
        //    gc_ID.Visible = true;
        //}

        //protected void ckDisplayIDColumns_Unchecked(object sender, RoutedEventArgs e)
        //{
        //    gc_ID.Visible = false;
        //}

        //protected void HideIDColumns(Nullable<bool> isChecked)
        //{
        //    if ((bool)isChecked)
        //    {
        //        gc_ID.Visible = true;
        //    }
        //    else
        //    {
        //        gc_ID.Visible = false;
        //    }
        //}

        protected void OnItemClick_Cmd1(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            MessageBox.Show(string.Format("{0}\n{1}",
                "What do you want to be able to do with this row",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));
        }
    }
}
