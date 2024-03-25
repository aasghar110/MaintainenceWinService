using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;

[RunInstaller(true)]
public class ProjectInstaller : Installer
{
    private ServiceInstaller serviceInstaller;
    private ServiceProcessInstaller processInstaller;

    public ProjectInstaller()
    {
        // Instantiate installer components
        serviceInstaller = new ServiceInstaller();
        processInstaller = new ServiceProcessInstaller();

        // Configure service installer
        serviceInstaller.ServiceName = "GEN_MaintenanceWinService"; // Set your service name here
        serviceInstaller.DisplayName = "GEN_MaintenanceWinService"; // Set your service display name here
        serviceInstaller.Description = "GEN_MaintenanceWinService: Monitors system performance, manages backups, and sends notifications for critical events."; // Set your service description here
        serviceInstaller.StartType = ServiceStartMode.Automatic;

        // Configure process installer
        processInstaller.Account = ServiceAccount.LocalSystem;

        // Add installers to collection
        Installers.Add(processInstaller);
        Installers.Add(serviceInstaller);

        // Start Service Automatically After Install
        // this.Committed += new InstallEventHandler(ProjectInstaller_Committed);
    }

    // private void ProjectInstaller_Committed(object sender, InstallEventArgs e)
    // {
    //     // Start the service after installation
    //     ServiceController sc = new ServiceController("GEN_MaintenanceWinService"); // Replace "GEN_MaintenanceWinService" with your service name
    //     sc.Start();
    // }
}
