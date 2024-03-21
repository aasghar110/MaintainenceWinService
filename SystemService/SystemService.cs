using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Management; 
using System.Net;
using System.ServiceProcess;
using System.Net.Http;
using System.Timers;
using System.Threading.Tasks;
using System.Runtime.Remoting.Messaging;
using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;
using System.Globalization;
using GTVeriSys_Net;
using System.Runtime.Remoting.Contexts;

namespace SystemService
{
    public partial class SystemService : ServiceBase
    {
        private Timer timer;
        private string isED;
        private string connectionString;
        private string CPUPercent;
        private string MemoryPercent;
        private string RestartHour;
        private string RestartMin;
        private string ZIPHour;
        private string ZIPMin;
        private string ZIP_IntervalHours;
        private string MobileNumbers;
        private HttpClient httpClient;
        private bool notificationSent = false;
        private string folderToZip;
        private string OutputZIP;
        private bool EnableZIPUpload;




        public SystemService()
        {
            InitializeComponent();
            isED = ConfigurationManager.AppSettings["isED"];
            if(isED == "Y")
            {
                string connectionStringEncoded = ConfigurationManager.AppSettings["ConnectionString"];

                connectionString = nDecodeGtVerySis(connectionStringEncoded);
            }
            else
            {
                connectionString = ConfigurationManager.AppSettings["ConnectionString"];
                nEncodeGtVerySis(connectionString);
            }
            httpClient = new HttpClient(); 

        }
        public string nDecodeGtVerySis(string DecodeWord)
        {
            try
            {
                GTVeriSys_Net.GTVS gtv = new GTVeriSys_Net.GTVS();
                var DecodeConn = gtv.nDecode(DecodeWord, "Y");
                return DecodeConn;
            }
            catch (Exception ex)
            {
                NotepadLog($"(nDecodeGtVerySis) {ex.Message} {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");
                return DecodeWord;
            }
        }
        public string nEncodeGtVerySis(string EncodeWord)
        {
            try
            {
                string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "EncodedConn.txt");
                GTVeriSys_Net.GTVS gtv = new GTVeriSys_Net.GTVS();
                var EncodedConn = gtv.nEncode(EncodeWord);
                using (StreamWriter writer = new StreamWriter(filePath, true))
                {
                    writer.WriteLine(EncodedConn);
                }

                return EncodedConn;
            }
            catch (Exception ex)
            {
                NotepadLog($"(nEncodeGtVerySis) {ex.Message} {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");
                return EncodeWord;
            }
        }

        protected async override void OnStart(string[] args)
        {
            notificationSent = false;
            string IpAddress = GetIpAddress();
            string ServerName = Environment.MachineName;
            timer = new Timer();
            timer.Interval = 50000; 
            timer.Elapsed += Timer_Elapsed;
            timer.AutoReset = true;
            timer.Start();
            float cpuUsage = GetCPUUsage();
            float memoryUsage = GetMemoryUsage();
            string LastBootTime = GetLastBootTime();
            DateTime SystemRestartDateTime = DateTime.Now;
            UpdateStartTimeFile();
            string Description = "Service Started";
            LogEvent(ServerName, IpAddress, cpuUsage.ToString(), memoryUsage.ToString(), Description,"3", LastBootTime);
            NotepadLog($"{Description} {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");

        }

        private async void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            try 
            {
                float cpuUsage = GetCPUUsage();
                float memoryUsage = GetMemoryUsage();
                string IpAddress = GetIpAddress();
                string ServerName = Environment.MachineName;




                EnableZIPUpload = Convert.ToBoolean(ConfigurationManager.AppSettings["EnableZIPUpload"]);
                ZIPHour = ConfigurationManager.AppSettings["ZIPHour"];
                int ZIPHourInt = Int32.Parse(ZIPHour);
                ZIPMin = ConfigurationManager.AppSettings["ZIPMin"];
                int ZIPMinInt = Int32.Parse(ZIPMin);
                ZIP_IntervalHours = ConfigurationManager.AppSettings["ZIP_CreationDays"];
                int ZIP_IntervalDays = Int32.Parse(ZIP_IntervalHours);
                DateTime startTime = GetStartTimeFromFile();

                if (EnableZIPUpload)
                {
                    if (startTime != DateTime.MinValue)
                    {
                        TimeSpan elapsedTime = DateTime.Now - startTime;

                        if (elapsedTime.TotalDays >= ZIP_IntervalDays)
                        {
                            if (DateTime.Now.Hour == ZIPHourInt && DateTime.Now.Minute == ZIPMinInt)
                            {
                                DateTime ZIPDateTime = DateTime.Now;
                                folderToZip = ConfigurationManager.AppSettings["ZIPPath"];
                                OutputZIP = ConfigurationManager.AppSettings["OutputZIP"];
                                string folderName = Path.GetFileName(folderToZip);
                                string DirectoryName = Path.GetDirectoryName(folderToZip);
                                string formattedZipDateTime = ZIPDateTime.ToString("ddMMyyyy");
                                string zipFileName = $"{folderName}{formattedZipDateTime}.zip";
                                CreateZipFile(folderToZip, zipFileName, OutputZIP);
                                UpdateStartTimeFile();

                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("EnableZIPUpload is False");
                    NotepadLog($"EnableZIPUpload is False {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");

                }




                RestartHour = ConfigurationManager.AppSettings["RestartHour"];
                int RestartHourInt = Int32.Parse(RestartHour);
                RestartMin = ConfigurationManager.AppSettings["RestartMin"];
                int RestartMinInt = Int32.Parse(RestartMin);
                MobileNumbers = ConfigurationManager.AppSettings["MobileNumbers"];

                if (DateTime.Now.Hour == RestartHourInt && DateTime.Now.Minute == RestartMinInt)
                {
                    // Log restart 
                    DateTime SystemRestartDateTime = DateTime.Now;
                    string formattedRestartDateTime = SystemRestartDateTime.ToString("dd-MM-yyyy HH:mm");
                    string Description = "Restart";
                    string LastBootTime = GetLastBootTime();
                    string messageRestart = $"Restart - {ServerName}, IP: {IpAddress}, CPU: {cpuUsage}%, Mem: {memoryUsage}%, Last: {LastBootTime}, Restart: {formattedRestartDateTime}";
                    await CallApi(messageRestart, MobileNumbers);
                    LogEvent(ServerName, IpAddress, cpuUsage.ToString(), memoryUsage.ToString(), Description, "0", LastBootTime);
                    NotepadLog($"{Description} {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");
                    notificationSent = false;
                    System.Diagnostics.Process.Start("shutdown", "/r /f /t 0");
                }



                CPUPercent = ConfigurationManager.AppSettings["CPUPercent"];
                int CPUPercentInt = Int32.Parse(CPUPercent);
                MemoryPercent = ConfigurationManager.AppSettings["MemoryPercent"];
                int MemoryPercentInt = Int32.Parse(MemoryPercent);



                if ((cpuUsage > CPUPercentInt || memoryUsage > MemoryPercentInt) && !notificationSent)
                {
                    string LastBootTime = GetLastBootTime();
                    DateTime SystemRestartDateTime = DateTime.Now;
                    string formattedRestartDateTime = SystemRestartDateTime.ToString("dd-MM-yyyy HH:mm");
                    string message = $"High Usage - {ServerName}, IP: {IpAddress}, CPU: {cpuUsage}%, Mem: {memoryUsage}%, Last: {LastBootTime}, Time: {formattedRestartDateTime}";
                    await CallApi(message, MobileNumbers);
                    string Description = "CPU OR Memory High";
                    LogEvent(ServerName, IpAddress, cpuUsage.ToString(), memoryUsage.ToString(), message, "1", LastBootTime);
                    NotepadLog($"{Description} {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");

                    notificationSent = true;


                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"(Timer_Elapsed) Error logging event: {ex.Message}");
                NotepadLog($"(Timer_Elapsed) {ex.Message} {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");

            }
            
        }

        protected override void OnStop()
        {
            timer.Stop();
            timer.Dispose();
            notificationSent = false;
            string IpAddress = GetIpAddress();
            string ServerName = Environment.MachineName;
            string LastBootTime = GetLastBootTime();
            float cpuUsage = GetCPUUsage();
            float memoryUsage = GetMemoryUsage();
            LogEvent(ServerName, IpAddress, cpuUsage.ToString(), memoryUsage.ToString(), "Service Stopped", "3", LastBootTime);
            NotepadLog($"Service Stopped {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");


        }

        private void LogEvent(string SystemName, string IpAddress, string CPU, string Memory, string Description, string LogType , string LastRestartDateTime= "") 
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand("GT_SystemLogSP", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@nType", 0);
                        command.Parameters.AddWithValue("@nstype", 0);
                        command.Parameters.AddWithValue("@SystemName", SystemName);
                        command.Parameters.AddWithValue("@IPAddress", IpAddress);
                        command.Parameters.AddWithValue("@LastRestartDateTime", LastRestartDateTime);
                        command.Parameters.AddWithValue("@CPU", CPU);
                        command.Parameters.AddWithValue("@Memory", Memory);
                        command.Parameters.AddWithValue("@Description", Description);
                        command.Parameters.AddWithValue("@Type", LogType);
                        command.Parameters.AddWithValue("@DataEntryUserid", $"Gen_SystemService({SystemName})");
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"(LogEvent) Error logging event: {ex.Message}");
                NotepadLog($"(LogEvent) {ex.Message} {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");

            }
        }

        private float GetCPUUsage()
        {

            PerformanceCounter cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
            cpuCounter.NextValue();
            System.Threading.Thread.Sleep(1000);
            float cpuUsage = cpuCounter.NextValue();
            cpuUsage = (float)Math.Round(cpuUsage, 2);
            return cpuUsage;

        }

        private float GetMemoryUsage()
        {

            ulong totalRAM = 0;

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
            foreach (ManagementObject obj in searcher.Get())
            {
                totalRAM = (ulong)obj["TotalPhysicalMemory"];
                break;
            }
            ulong TotalSystemRAM = totalRAM / (1024 * 1024);

            var memoryCounter = new PerformanceCounter("Memory", "Available MBytes");
            float availableMemory = memoryCounter.NextValue();
            float memoryUsage = ((TotalSystemRAM - availableMemory) / TotalSystemRAM) * 100;
            memoryUsage = (float)Math.Round(memoryUsage, 2);
            return memoryUsage;
        }
        private string GetIpAddress()
        {
            string ipAddress = "";

            try
            {
                IPAddress[] addresses = Dns.GetHostEntry(Dns.GetHostName()).AddressList;
                foreach (IPAddress address in addresses)
                {
                    if (address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                    {
                        ipAddress = address.ToString();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                return $"error: {ex}";
            }
            return ipAddress;

        }
        private string GetLastBootTime()
        {
            try
            {
                ManagementObjectSearcher mos = new ManagementObjectSearcher("SELECT LastBootUpTime FROM Win32_OperatingSystem WHERE Primary='true'");

                foreach (ManagementObject mo in mos.Get())
                {
                    DateTime lastBootTime = ManagementDateTimeConverter.ToDateTime(mo["LastBootUpTime"].ToString());

                    string formattedBootTime = lastBootTime.ToString("yyyy-MM-dd HH:mm:ss");


                    return formattedBootTime; 
                }
            }
            catch (Exception ex)
            {
                return $"error: {ex}";
            }

            return "Unable to retrieve last boot time.";
        }
        private async Task<string> CallApi(string message, string MobileNumbers)
        {
            try
            {
                bool enableSMSApi = Convert.ToBoolean(ConfigurationManager.AppSettings["EnableSMSAPI"]);

                if (enableSMSApi)
                {
                    string apiUrl = $"https://pk.eocean.net//APIManagement/API/RequestAPI?user=GENTEC&pwd=ADcCywfmYRaXL20wWMJTbiEOJDhWoxhdsfciOhR3D6GSSyM1OiPsEPmuw6C%2f4%2bIAfw%3d%3d&sender=GENTEC&reciever={MobileNumbers}&msg-data={message}&response=string";

                    HttpResponseMessage response = await httpClient.GetAsync(apiUrl);

                    if (response.IsSuccessStatusCode)
                    {
                        string responseData = await response.Content.ReadAsStringAsync();
                        return responseData;
                    }
                    else
                    {
                        Console.WriteLine($"API call failed with status code: {response.StatusCode}");
                        NotepadLog($"API call failed with status code: {response.StatusCode}");

                        return null;
                    }
                }
                else
                {
                    Console.WriteLine("EnableSMSAPI is False");
                    NotepadLog($"EnableSMSAPI is False {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");

                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error calling API: {ex.Message}");
                NotepadLog($"(CallApi) {ex.Message} {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");
                return null;
            }
        }


        private async void CreateZipFile(string folderToZip, string zipFileName, string OutputPath)
        {
            try
            {
                if (!Directory.Exists(folderToZip))
                {
                    Console.WriteLine("Folder to zip does not exist.");
                    NotepadLog($"Folder to zip does not exist. {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");
                    return;
                }

                string parentDirectory = Directory.GetParent(folderToZip).FullName;
                string destinationFolder = parentDirectory;
                string zipFilePath = Path.Combine(OutputPath, zipFileName);

                if (File.Exists(zipFilePath))
                {
                    File.Delete(zipFilePath); // Delete existing zip file
                }

                using (ZipOutputStream zipStream = new ZipOutputStream(File.Create(zipFilePath)))
                {
                    zipStream.SetLevel(9); // Set the compression level

                    byte[] buffer = new byte[4096];

                    // Recursively add files to the zip
                    string[] files = Directory.GetFiles(folderToZip, "*", SearchOption.AllDirectories);
                    foreach (string file in files)
                    {
                        string relativePath = GetRelativePath(folderToZip, file); // Calculate relative path
                        relativePath = relativePath.Replace('\\', '/'); // Ensure Unix-style path separators for zip
                        ZipEntry entry = new ZipEntry(relativePath);
                        zipStream.PutNextEntry(entry);

                        using (FileStream fs = File.OpenRead(file))
                        {
                            StreamUtils.Copy(fs, zipStream, buffer);
                        }

                        zipStream.CloseEntry();
                    }
                }

                await UploadZipFile(zipFilePath);


            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating zip file: {ex.Message}");
                NotepadLog($"(CreateZipFile) Error creating zip file: {ex.Message} {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");

            }
        }

        private string GetRelativePath(string basePath, string fullPath)
        {
            Uri baseUri = new Uri(basePath);
            Uri fullUri = new Uri(fullPath);
            return baseUri.MakeRelativeUri(fullUri).ToString();
        }

        private async Task UploadZipFile(string zipFilePath)
        {
            try
            {
                string ftpServer = ConfigurationManager.AppSettings["FtpServer"];
                string ftpUsername = ConfigurationManager.AppSettings["FtpUsername"];
                string ftpPassword = ConfigurationManager.AppSettings["FtpPassword"];
                string ftpPort = ConfigurationManager.AppSettings["FtpPort"];

                string fileName = Path.GetFileName(zipFilePath);
                string ftpUri = $"{ftpServer}:{ftpPort}/{fileName}";

                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(ftpUri);
                request.Method = WebRequestMethods.Ftp.UploadFile;
                request.Credentials = new NetworkCredential(ftpUsername, ftpPassword);

                using (Stream fileStream = File.OpenRead(zipFilePath))
                using (Stream ftpStream = await request.GetRequestStreamAsync())
                {
                    byte[] buffer = new byte[8192]; 
                    int bytesRead;
                    while ((bytesRead = await fileStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                    {
                        await ftpStream.WriteAsync(buffer, 0, bytesRead);
                    }
                }

                using (FtpWebResponse response = (FtpWebResponse)await request.GetResponseAsync())
                {
                    Console.WriteLine($"Upload ZIP File Complete, status {response.StatusDescription}");
                    NotepadLog($"Upload ZIP File Complete, status {response.StatusDescription} {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");


                }
                File.Delete(zipFilePath);
                await ZipUploadedSuccessfully(zipFilePath);

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error uploading ZIP file: {ex.Message}");
                NotepadLog($"(UploadZipFile) {ex.Message} {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");

            }
        }
        private async Task ZipUploadedSuccessfully(string zipFilePath)
        {
            try
            {
                string IpAddress = GetIpAddress();
                string ServerName = Environment.MachineName;
                timer = new Timer();
                timer.Interval = 50000;
                timer.Elapsed += Timer_Elapsed;
                timer.AutoReset = true;
                timer.Start();
                float cpuUsage = GetCPUUsage();
                float memoryUsage = GetMemoryUsage();
                string LastBootTime = GetLastBootTime();
                DateTime SystemRestartDateTime = DateTime.Now;
                UpdateStartTimeFile();
                string message = $"{Path.GetFileName(zipFilePath)} file created and uploaded to FTP successfully";
                LogEvent(ServerName, IpAddress, cpuUsage.ToString(), memoryUsage.ToString(), message, "2", LastBootTime);
                NotepadLog($"{message} {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");
                await CallApi(message, MobileNumbers);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error calling API after ZIP upload: {ex.Message}");
                NotepadLog($"(CallApiAfterZipUpload) Error calling API after ZIP upload: {ex.Message} {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");

            }
        }
        private void UpdateStartTimeFile()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ServiceStartTime.txt");

            try
            {
                if (File.Exists(filePath))
                {
                    using (StreamWriter writer = new StreamWriter(filePath))
                    {
                        writer.WriteLine($"Service started at: {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");
                    }
                }
                else
                {
                    using (StreamWriter writer = File.CreateText(filePath))
                    {
                        writer.WriteLine($"Service started at: {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating start time file: {ex.Message}");
                NotepadLog($"(UpdateStartTimeFile) Error updating start time file: {ex.Message} {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");

            }
        }

        private DateTime GetStartTimeFromFile()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ServiceStartTime.txt");

            try
            {
                if (File.Exists(filePath))
                {
                    string content = File.ReadAllText(filePath);

                    string[] lines = content.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                    if (lines.Length > 0)
                    {
                        string timeString = lines[0].Replace("Service started at: ", "");

                        if (DateTime.TryParseExact(timeString, "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startTime))
                        {
                            return startTime;
                        }
                        else
                        {
                            Console.WriteLine("Failed to parse start time.");
                            return DateTime.MinValue;
                        }
                    }
                    else
                    {
                        Console.WriteLine("File is empty.");
                        return DateTime.MinValue;
                    }
                }
                else
                {
                    Console.WriteLine("Start time file not found.");
                    return DateTime.MinValue;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading start time file: {ex.Message}");
                NotepadLog($"(GetStartTimeFromFile) Error reading start time file: {ex.Message} {DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")}");

                return DateTime.MinValue;
            }
        }

        private void NotepadLog(string LogText)
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SystemServiceLog.txt");

            try
            {
                using (StreamWriter writer = new StreamWriter(filePath, true)) 
                {
                    writer.WriteLine(LogText);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

    }
}

