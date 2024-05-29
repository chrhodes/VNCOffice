using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class ConnectionPointRowWrapper : ModelWrapper<ConnectionPointRow>
    {
        public ConnectionPointRowWrapper(ConnectionPointRow model) : base(model)
        {
        }

        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
        public string X { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Y { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DirX { get { return GetValue<string>(); } set { SetValue(value); } }
        public string A { get { return GetValue<string>(); } set { SetValue(value); } }
        public string DirY { get { return GetValue<string>(); } set { SetValue(value); } }
        public string B { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Type { get { return GetValue<string>(); } set { SetValue(value); } }
        public string C { get { return GetValue<string>(); } set { SetValue(value); } }
        public string D { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
