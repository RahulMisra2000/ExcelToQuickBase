using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;

namespace CustomActionProject
{
    [RunInstaller(true)]
    public partial class CustomAction : System.Configuration.Install.Installer
    {
        public CustomAction()
        {
            InitializeComponent();
        }

        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);
        }

        /// <summary>
        /// This is called by the Setup Project immediately after a successful installation
        /// </summary>
        /// <param name="savedState"></param>
        public override void Commit(IDictionary savedState)
        {
            base.Commit(savedState);
            // The Setup Project populates the installation folder location (chosen by the user) into the abcd parameter
            string installationFolder = this.Context.Parameters["abcd"];            

            // Calling an executable that will encrypt the sensitive stuf inside the .config file
            ProcessStartInfo si = new ProcessStartInfo(@"c:\temp\ReadExcelProject.exe", "-sp:-encOnly" + " -if:" + installationFolder);
            si.WindowStyle = ProcessWindowStyle.Hidden;

            Process p;

            try
            {                
                p = Process.Start(si);                
                p.WaitForExit();                
            }
            catch (Exception ex)
            {
                Context.LogMessage("Failed to Eccrypt " + ex);
            }
            
        }

        public override void Rollback(IDictionary savedState)
        {
            base.Rollback(savedState);
        }

        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);
        }
    }
}
