using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace SupportTools_Excel.User_Interface
{
    public class Help
    {
        public static void GetHelpOnTopic(string helpTopic)
        {

            switch (helpTopic)
            {
                case "DatabaseExpandTemplate":
                    MessageBox.Show("Display Help On " + helpTopic);
                    break;

                case "InstanceExpandTemplate":
                    MessageBox.Show("Display Help On " + helpTopic);
                    break;

                case "Servers":
                    MessageBox.Show("Display Help On " + helpTopic);
                    break;

                case "SnapShotControlsTemplate":
                    MessageBox.Show("Display Help On " + helpTopic);
                    break;

                default:
                    MessageBox.Show("Default Display Help On " + helpTopic);
                    break;
            }

        }

    }
}
