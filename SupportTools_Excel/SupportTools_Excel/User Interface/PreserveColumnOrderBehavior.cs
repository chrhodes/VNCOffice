using System.Linq;
using System.Windows;

using DevExpress.Xpf.Grid;

using Microsoft.Xaml.Behaviors;

namespace SupportTools_Excel.User_Interface
{
    public class PreserveColumnOrderBehavior : Behavior<GridControl>
    {
        int[] visibilities = null;

        public object ChangeDetector
        {
            get { return (object)GetValue(ChangeDetectorProperty); }
            set { SetValue(ChangeDetectorProperty, value); }
        }

        public static readonly DependencyProperty ChangeDetectorProperty =
             DependencyProperty.Register(
                "ChangeDetector", 
                typeof(object), 
                typeof(PreserveColumnOrderBehavior), 
                new PropertyMetadata(null, ChangeDetected));

        public static void ChangeDetected(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var me = d as PreserveColumnOrderBehavior;

            GridControl gridControl = me.AssociatedObject;

            if (gridControl == null || gridControl.Columns.Count == 0) return;

            if (me.visibilities == null)
            {
                // initial call
                me.visibilities = gridControl.Columns.Select(x => x.VisibleIndex).ToArray();
            }
            else
            {
                for (int i = 0; i < gridControl.Columns.Count; i++)
                    if (gridControl.Columns[i].VisibleIndex != me.visibilities[i])
                        gridControl.Columns[i].VisibleIndex = me.visibilities[i];
            }
        }

    }
}
