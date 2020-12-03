using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Threading.Tasks;

namespace AutoGenEpoService
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
        }

        private void MtkEpoGenServiceInstaller_AfterInstall(object sender, InstallEventArgs e)
        {

        }

        private void MtkEpoGenServiceProcessInstaller_AfterInstall(object sender, InstallEventArgs e)
        {

        }
    }
}
