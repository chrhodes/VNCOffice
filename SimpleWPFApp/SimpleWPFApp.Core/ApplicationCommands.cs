using Prism.Commands;

namespace SimpleWPFApp.Core
{
    public class ApplicationCommands
    {
        public static CompositeCommand NavigateCommand = new CompositeCommand();
        public static CompositeCommand OpenShellCommand = new CompositeCommand();
    }
}
