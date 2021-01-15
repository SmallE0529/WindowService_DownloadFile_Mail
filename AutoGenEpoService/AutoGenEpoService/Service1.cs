using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net.Http;
using System.Timers;
using MimeKit;

namespace AutoGenEpoService
{
    public partial class Service1 : ServiceBase
    {
        private EventLog eventLog1;
        private Timer mTimer;
        private static int min = 0;

        private static bool InternetConnected { get; set; }

        /* File Define */
        private static string path = @"C:\Folder\mLog.log";
        private static string ConfigFile = @"C:\Folder\mConfig.cfg";
        private static string BatFile = @"C:\Folder\BatFileName.bat";
        private static string HexFile = @"C:\Folder\HexFileName.hex";
        private static string BinFile = @"C:\Folder\BinFileName.bin";
        private static string InfoFile = @"C:\Folder\mInfo.txt";

        private static readonly HttpClient httpClient = new HttpClient();
        private static HttpResponseMessage response;

        /* URL */
        private static String URL = "http://www.google.com";

        public Service1()
        {
            InitializeComponent();

            this.AutoLog = false;
            eventLog1 = new EventLog();

            if (!System.Diagnostics.EventLog.SourceExists("ServiceSource"))
            {

                System.Diagnostics.EventLog.CreateEventSource("ServiceSource", "ServiceLog");

            }

            eventLog1.Source = "ServiceSource";
            eventLog1.Log = "ServiceLog";
        }

        protected override void OnStart(string[] args)
        {
            mTimer = new Timer();
            mTimer.Elapsed += new ElapsedEventHandler(mTimer_Elapsed);
            mTimer.Interval = 60 * 1000; // uint: ms
            mTimer.Start();
            eventLog1.WriteEntry("Download File Timer Start", EventLogEntryType.Information);
        }

        private void mTimer_Elapsed(object sender, EventArgs e)
        {
            min++;
            if (min >= 1) // 1 min check file
            {
                min = 0;
                RunTask();
            }
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("Download File Timer Stop", EventLogEntryType.Information);
            mTimer.Stop();
            mTimer = null;
        }

        public static void RunTask()
        {
            DateTime FileBuildTime = File.GetLastWriteTime(BinFile);
            DateTime now = DateTime.Now;
            TimeSpan diff = DateTime.Now.Subtract(FileBuildTime);

            if ((diff.TotalHours <= 12.0) || (now.Hour < 8))
            {
                return;
            }

            InternetConnected = System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable();

            if (InternetConnected)
            {
                if (Download())
                {
                    GenBinFile();
                    UpdateInfo();
                }
            }
            else
            {
                File.AppendAllText(path, DateTime.Now + ":\t" + "Internet UnConnect\r\n");
            }
            SentMail();
        }

        public static void GenBinFile()
        {
            ProcessStartInfo mAction = new ProcessStartInfo();
            mAction.FileName = "BatFileName.bat"; // exec file name
            mAction.WorkingDirectory = @"C:\Folder";
            Process star = Process.Start(mAction);

            if (star.Start())
            {
                star.WaitForExit();
            }

            star.Close();
        }

        public static bool Download()
        {
            var sDownloadUrl = URL;

            response = httpClient.GetAsync(new Uri(sDownloadUrl)).Result;
            if (response.IsSuccessStatusCode)
            {
                var fileStream = File.Create(HexFile);
                var EpoStream = response.Content.ReadAsStreamAsync().Result;
                EpoStream.Seek(0, SeekOrigin.Begin);
                EpoStream.CopyTo(fileStream);
                fileStream.Close();
                return true;
            }
            File.AppendAllText(path, DateTime.Now + ":\t" + "Download Web Error\r\n");
            return false;
        }

        public static void UpdateInfo()
        {
            if (!File.Exists(InfoFile))
            {
                File.AppendAllText(path, DateTime.Now + ":\t" + "Info File not Exists\r\n");
                return;
            }

            if (!File.Exists(HexFile))
            {
                File.AppendAllText(path, DateTime.Now + ":\t" + "Hex File File not Exists\r\n");
                return;
            }

            if (!File.Exists(BinFile))
            {
                File.AppendAllText(path, DateTime.Now + ":\t" + "Bin File File not Exists\r\n");
                return;
            }

            string strCompare = "\"Info\":{\"Ver\":\"";
            var HexFileData = File.ReadAllBytes(HexFile);
            var BinFileData = File.ReadAllBytes(BinFile);
            var InfoData = File.ReadAllLines(InfoFile);

            UInt32 mTime = 0;
            UInt32 mWeek = 0;
            UInt32 mTow = 0;

            mTime = HexFileData[2];
            mTime = (mTime << 8) + HexFileData[1];
            mTime = (mTime << 8) + HexFileData[0];
            mWeek = mTime * 3600 / 604800;
            mTow = mTime * 3600 % 604800;


            String verInfo = 12.ToString("X") + "." + mWeek.ToString("X") + "." + mTow.ToString("X");

            UInt32 BinFileCrc32 = CalculateCrc32(BinFileData, 0);
            byte[] b_BinFileCrc32 = BitConverter.GetBytes(BinFileCrc32);
            Array.Reverse(b_BinFileCrc32);
            String sBinFileCrc32 = BitConverter.ToString(b_BinFileCrc32).Replace("-", "");

            int strLine = 0;

            for (int i = 0; i < InfoData.Length; i++)
            {
                if (InfoData[i].Contains(strCompare))
                {
                    strLine = i;
                    break;
                }
            }

            string strInfo = " " + "\"Info\"" + ":{" + "\"Ver\":" + "\"" + verInfo + "\"" + ",\"File\":\"BinFile.bin\"," + "\"CRC\":" + "\"" + sBinFileCrc32 + "\"" + "},";
            InfoData[strLine] = strInfo;
            File.WriteAllLines(InfoFile, InfoData);
        }

        static public UInt32 CalculateCrc32(byte[] buf, UInt32 crc)
        {
            UInt32[] lut = new uint[] {
                0x00000000, 0x1DB71064, 0x3B6E20C8, 0x26D930AC, 0x76DC4190, 0x6B6B51F4, 0x4DB26158, 0x5005713C,
                0xEDB88320, 0xF00F9344, 0xD6D6A3E8, 0xCB61B38C, 0x9B64C2B0, 0x86D3D2D4, 0xA00AE278, 0xBDBDF21C
            };
            crc = ~crc;
            foreach (UInt32 b in buf)
            {
                crc = lut[(crc ^ b) & 0x0F] ^ (crc >> 4);
                crc = lut[(crc ^ (b >> 4)) & 0x0F] ^ (crc >> 4);
            }
            return ~crc;
        }

        static public void SentMail()
        {
            if (!File.Exists(ConfigFile))
            {
                File.AppendAllText(path, DateTime.Now + ":\t" + "Cep Config File Not Exists\r\n");
                return;
            }
            // set congif
            var cfg = File.ReadAllLines(ConfigFile);

            MimeMessage message = new MimeMessage();
            message.From.Add(new MailboxAddress(cfg[0], cfg[1]));
            message.To.Add(new MailboxAddress("", cfg[1]));


            for (int i = 3; i < cfg.Length; i++)
            {
                message.To.Add(new MailboxAddress("", cfg[i]));
            }

            message.Subject = "Mail Subject";

            var builder = new BodyBuilder();

            builder.TextBody = @"Hi,Mail Body";
            builder.Attachments.Add(path);
            builder.Attachments.Add(BinFile);

            message.Body = builder.ToMessageBody();

            var smtp = new MailKit.Net.Smtp.SmtpClient();
            smtp.MessageSent += (sender, args) =>
            { // args.Response 
            };

            smtp.ServerCertificateValidationCallback = (s, c, h, e) => true;

            smtp.Connect("host", 0, MailKit.Security.SecureSocketOptions.StartTls);

            smtp.Authenticate(cfg[0], cfg[2]);

            try
            {
                smtp.Send(message);
                smtp.Disconnect(true);
            }
            catch (Exception ex)
            {
                File.AppendAllText(path, DateTime.Now + ":\t" + "Send e-mail Fail\r\n");
            }
        }
    }
}
