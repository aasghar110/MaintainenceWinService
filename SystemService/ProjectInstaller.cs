using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;

[RunInstaller(true)]
public class ProjectInstaller : Installer
{
    private ServiceInstaller serviceInstaller;
    private ServiceProcessInstaller serviceProcessInstaller1;
    private ServiceInstaller serviceInstaller1;
    private ServiceProcessInstaller processInstaller;

    public ProjectInstaller()
    {
        // Instantiate installer components
        serviceInstaller = new ServiceInstaller();
        processInstaller = new ServiceProcessInstaller();

        // Configure service installer
        serviceInstaller.ServiceName = "Gen_System_Service"; // Set your service name here
        serviceInstaller.StartType = ServiceStartMode.Automatic;

        // Configure process installer
        processInstaller.Account = ServiceAccount.LocalSystem;

        // Add installers to collection
        Installers.Add(processInstaller);
        Installers.Add(serviceInstaller);
    }

    private void InitializeComponent()
    {
            this.serviceProcessInstaller1 = new System.ServiceProcess.ServiceProcessInstaller();
            this.serviceInstaller1 = new System.ServiceProcess.ServiceInstaller();
            // 
            // serviceProcessInstaller1
            // 
            this.serviceProcessInstaller1.Account = System.ServiceProcess.ServiceAccount.LocalService;
            this.serviceProcessInstaller1.Password = null;
            this.serviceProcessInstaller1.Username = null;
            // 
            // serviceInstaller1
            // 
            this.serviceInstaller1.Description = "Gen_System_Service";
            this.serviceInstaller1.DisplayName = "Gen_System_Service";
            this.serviceInstaller1.ServiceName = "Gen_System_Service";
            this.serviceInstaller1.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.serviceProcessInstaller1,
            this.serviceInstaller1});

    }
}
