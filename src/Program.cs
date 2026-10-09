using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

[assembly: AssemblyTitle("Letter Merger")]
[assembly: AssemblyDescription("Local CSV and photo merge into the supplied A7 student thank-you card template")]
[assembly: AssemblyProduct("Letter Merger")]
[assembly: AssemblyCompany("Kevin Le")]
[assembly: AssemblyCopyright("Kevin Le")]
[assembly: AssemblyVersion("16.1.0.0")]
[assembly: AssemblyFileVersion("16.1.0.0")]
namespace LetterMerger
{
    public static class App
    {
        public static readonly string Root = AppDomain.CurrentDomain.BaseDirectory;
        public static readonly string LogPath = Path.Combine(Root, "A7_Startup.log");
        public static void Log(string message)
        {
            try
            {
                Core.RejectLink(LogPath);
                File.AppendAllText(LogPath, DateTime.Now.ToString("s") + " " + message + Environment.NewLine);
            }
            catch
            {
            }
        }

        public static void LogException(Exception error)
        {
            Log("ErrorType=" + error.GetType().FullName + " Code=" + error.HResult.ToString("X8"));
        }

        [STAThread]
        public static int Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                Application.ThreadException += (sender, e) =>
                {
                    LogException(e.Exception);
                    MessageBox.Show(e.Exception.GetBaseException().Message, "Letter Merger", MessageBoxButtons.OK, MessageBoxIcon.Error);
                };
                Log("Letter Merger v16 managed application starting.");
                using (var form = new MainForm(Root))
                {
                    Application.Run(form);
                }

                Log("Application closed normally.");
                return 0;
            }
            catch (Exception e)
            {
                LogException(e);
                try
                {
                    MessageBox.Show("The application could not start. " + e.GetBaseException().Message + Environment.NewLine + "Extract the full application folder to a location you can write to." + Environment.NewLine + "Log: " + LogPath, "Letter Merger", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch
                {
                }

                return 1;
            }
        }
    }
}
