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
        
        // for Test
        private static string path = @"C:\updateWithDount\VivoWatchBPtest\fw\MtkGpsTool\AutoGetMtkEpo.log";
        private static string ConfigFile = @"C:\updateWithDount\VivoWatchBPtest\fw\MtkGpsTool\EpoConfig.cfg";
        private static string MtkGpsBat = @"C:\updateWithDount\VivoWatchBPtest\fw\MtkGpsTool\GenVioletVfw.bat";
        private static string MtkGpsEpoFile = @"C:\updateWithDount\VivoWatchBPtest\fw\MtkGpsTool\EPO_GR_3_1.DAT";
        private static string MtkGpsEpoVfw = @"C:\updateWithDount\VivoWatchBPtest\fw\EPO_GR_3_1.vfw";
        private static string FwInfoFile = @"C:\updateWithDount\VivoWatchBPtest\fw\FW_info.txt";

        /*
        // for Online
        private static string path = @"C:\updateWithDount\Violet\fw\MtkGpsTool\AutoGetMtkEpo.log";
        private static string ConfigFile = @"C:\updateWithDount\Violet\fw\MtkGpsTool\EpoConfig.cfg";
        private static string MtkGpsBat = @"C:\updateWithDount\Violet\fw\MtkGpsTool\GenVioletVfw.bat";
        private static string MtkGpsEpoFile = @"C:\updateWithDount\Violet\fw\MtkGpsTool\EPO_GR_3_1.DAT";
        private static string MtkGpsEpoVfw = @"C:\updateWithDount\Violet\fw\EPO_GR_3_1.vfw";
        private static string FwInfoFile = @"C:\updateWithDount\Violet\fw\FW_info.txt";
        */

        //private static string FwInfoBetaFile = @"C:\updateWithDount\VivoWatchBP\fw\FW_info_Beta.txt";

        private static readonly HttpClient httpClient = new HttpClient();
        private static HttpResponseMessage response;

        /* 3 Day */
        private static String MTK_EPO_URL = "http://wpepodownload.mediatek.com/EPO_GR_3_1.DAT?vendor=ASUS&project=ASUS&device_id=EeDaFACZwOdplaFA-5t2J4Qb-gH9wt1TRDmudqI3SbM";
        /* 7 Day */
        //private static String MTK_EPO_URL = "http://epodownload.mediatek.com/EPO.DAT";

        public Service1()
        {
            InitializeComponent();

            this.AutoLog = false;
            eventLog1 = new EventLog();

            if (!System.Diagnostics.EventLog.SourceExists("MtkEpoSource"))
            {

                System.Diagnostics.EventLog.CreateEventSource("MtkEpoSource", "MtkEpoLog");

            }

            eventLog1.Source = "MtkEpoSource";
            eventLog1.Log = "MtkEpoLog";
        }

        protected override void OnStart(string[] args)
        {
            mTimer = new Timer();
            mTimer.Elapsed += new ElapsedEventHandler(mTimer_Elapsed);
            mTimer.Interval = 60 * 1000;
            mTimer.Start();
            eventLog1.WriteEntry("Gen MTK EPO Timer Start", EventLogEntryType.Information);
        }

        private void mTimer_Elapsed(object sender, EventArgs e)
        {
            min++;
            if (min >= 20) // real: 20 test: 1
            {
                min = 0;
                GenEpoRun();
            }
        }

        protected override void OnStop()
        {
            eventLog1.WriteEntry("Gen MTK EPO Timer Stop", EventLogEntryType.Information);
            mTimer.Stop();
            mTimer = null;
        }

        public static void GenEpoRun()
        {
            DateTime FileBuildTime = File.GetLastWriteTime(MtkGpsEpoVfw);
            DateTime now = DateTime.Now;
            TimeSpan diff = DateTime.Now.Subtract(FileBuildTime);

            if ((diff.TotalHours <= 16.0) || (now.Hour < 9))
            {
                return;
            }

            InternetConnected = System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable();

            if (InternetConnected)
            {
                if (Download())
                {
                    GenEpoVfw();
                    UpdateEpoInfo();
                }
            }
            else
            {
                File.AppendAllText(path, DateTime.Now + ":\t" + "Internet UnConnect\r\n");
            }
            SentMail();
        }

        public static void GenEpoVfw()
        {
            ProcessStartInfo ActionGenEpo = new ProcessStartInfo();
            ActionGenEpo.FileName = "GenVioletVfw.bat";//執行的檔案名稱
            ActionGenEpo.WorkingDirectory = @"C:\updateWithDount\VivoWatchBPtest\fw\MtkGpsTool";
            Process star = Process.Start(ActionGenEpo);

            if (star.Start())
            {
                star.WaitForExit();
            }

            star.Close();
        }

        public static bool Download()
        {
            var sDownloadUrl = MTK_EPO_URL;

            response = httpClient.GetAsync(new Uri(sDownloadUrl)).Result;
            if (response.IsSuccessStatusCode)
            {
                var fileStream = File.Create(MtkGpsEpoFile);
                var EpoStream = response.Content.ReadAsStreamAsync().Result;
                EpoStream.Seek(0, SeekOrigin.Begin);
                EpoStream.CopyTo(fileStream);
                fileStream.Close();
                return true;
            }
            File.AppendAllText(path, DateTime.Now + ":\t" + "MTK Download Web Error\r\n");
            return false;
        }

        public static void UpdateEpoInfo()
        {
            if (!File.Exists(FwInfoFile))
            {
                File.AppendAllText(path, DateTime.Now + ":\t" + "FW_info File not Exists\r\n");
                return;
            }

            if (!File.Exists(MtkGpsEpoFile))
            {
                File.AppendAllText(path, DateTime.Now + ":\t" + "EPO_GR_3_1 DAT File not Exists\r\n");
                return;
            }

            if (!File.Exists(MtkGpsEpoVfw))
            {
                File.AppendAllText(path, DateTime.Now + ":\t" + "EPO_GR_3_1 vfw File not Exists\r\n");
                return;
            }

            string strGpsCompare = "\"GPS\":{\"Ver\":\"";
            var MtkGpsEpoData = File.ReadAllBytes(MtkGpsEpoFile);
            var MtkGpsVfwData = File.ReadAllBytes(MtkGpsEpoVfw);
            var FwInfoData = File.ReadAllLines(FwInfoFile);

            UInt32 GPS_Time = 0;
            UInt32 GPS_Week = 0;
            UInt32 GPS_Tow = 0;

            GPS_Time = MtkGpsEpoData[2];
            GPS_Time = (GPS_Time << 8) + MtkGpsEpoData[1];
            GPS_Time = (GPS_Time << 8) + MtkGpsEpoData[0];
            GPS_Week = GPS_Time * 3600 / 604800;
            GPS_Tow = GPS_Time * 3600 % 604800;

            /*byte[] b_12 = BitConverter.GetBytes(12);
            byte[] b_GPS_Week = BitConverter.GetBytes(GPS_Week);
            byte[] b_GPS_Tow = BitConverter.GetBytes(GPS_Tow);
            
            Array.Reverse(b_12);
            Array.Reverse(b_GPS_Week);
            Array.Reverse(b_GPS_Tow);*/

            String EpoVerInfo = 12.ToString("X") + "." + GPS_Week.ToString("X") + "." + GPS_Tow.ToString("X");

            UInt32 BinFileCrc32 = CalculateCrc32(MtkGpsVfwData, 0);
            byte[] b_BinFileCrc32 = BitConverter.GetBytes(BinFileCrc32);
            Array.Reverse(b_BinFileCrc32);
            String sBinFileCrc32 = BitConverter.ToString(b_BinFileCrc32).Replace("-", "");

            int GpsLine = 0;

            for (int i = 0; i < FwInfoData.Length; i++)
            {
                if (FwInfoData[i].Contains(strGpsCompare))
                {
                    GpsLine = i;
                    break;
                }
            }

            string GpsFwInfo = " " + "\"GPS\"" + ":{" + "\"Ver\":" + "\"" + EpoVerInfo + "\"" + ",\"File\":\"EPO_GR_3_1.vfw\"," + "\"CRC\":" + "\"" + sBinFileCrc32 + "\"" + "},";
            FwInfoData[GpsLine] = GpsFwInfo;
            File.WriteAllLines(FwInfoFile, FwInfoData);
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
            message.To.Add(new MailboxAddress(cfg[0], cfg[1]));


            for (int i = 3; i < cfg.Length; i++)
            {
                message.To.Add(new MailboxAddress(cfg[0], cfg[i]));
            }

            message.Subject = "MTK GPS EPO Daily Report";

            var builder = new BodyBuilder();

            builder.TextBody = @"Hi,

MTK GPS EPO Daily Report

From violet.fw server

*** This is an automatically generated email, please do not reply***
";
            builder.Attachments.Add(path);
            builder.Attachments.Add(MtkGpsEpoVfw);

            message.Body = builder.ToMessageBody();

            var smtp = new MailKit.Net.Smtp.SmtpClient();
            smtp.MessageSent += (sender, args) =>
            { // args.Response 
            };

            smtp.ServerCertificateValidationCallback = (s, c, h, e) => true;

            smtp.Connect("mymail.asus.com", 25, MailKit.Security.SecureSocketOptions.StartTls);

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
