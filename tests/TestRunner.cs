using System;
using System.IO;
using System.Windows.Forms;

namespace LetterMerger
{
    public static class TestRunner
    {
        [STAThread]
        public static int Main(string[] args)
        {
            try
            {
                if (args.Length < 2)
                    throw new ArgumentException("Usage: LetterMerger.Tests.exe --core|--ui APP_ROOT [OUTPUT_DIRECTORY]");
                string root = Path.GetFullPath(args[1]);
                string output = args.Length > 2 ? Path.GetFullPath(args[2]) : Path.Combine(root, "Output", "SelfCheck_" + Guid.NewGuid().ToString("N"));
                if (args[0] == "--core")
                    Tests.Run(root, output);
                else if (args[0] == "--ui")
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Tests.UI(root, output);
                }
                else
                    throw new ArgumentException("Choose --core or --ui.");
                return 0;
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e);
                return 1;
            }
        }
    }
}
