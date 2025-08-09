using Prism.Events;

namespace SupportTools_Visio.Core
{
    public class SelectionChangedEvent : PubSubEvent { }
    public class LoadPageEvent : PubSubEvent { }
    public class SavePageEvent : PubSubEvent { }
    public class UseLinqToExcelEvent : PubSubEvent { }
    public class LoadExcelTableEvent : PubSubEvent { }

    public class UseExcelDataReaderEvent : PubSubEvent { }
    public class LoadExcelFileEvent : PubSubEvent { }
    public class ExecuteEvent : PubSubEvent { }
    public class ReloadXmlEvent : PubSubEvent { }

}
