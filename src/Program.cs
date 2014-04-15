using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Windows.Forms;

namespace bground
{
    class Program
    {
        private static readonly NotifyIcon TrayIcon = new NotifyIcon();
        private static readonly ContextMenu ContextMenu = new ContextMenu();
        private static readonly Timer Timer = new Timer();
        private static readonly IntPtr ThisConsole = GetConsoleWindow();

        [DllImport("kernel32.dll", ExactSpelling = true)]
        private static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SystemParametersInfo(uint uiAction, uint uiParam, string pvParam, uint fWinIni);

        [DllImport("Kernel32")]
        private static extern bool SetConsoleCtrlHandler(EventHandler handler, bool add);
        private delegate bool EventHandler(CtrlType sig);

        enum CtrlType
        {
            CTRL_C_EVENT = 0,
            CTRL_BREAK_EVENT = 1,
            CTRL_CLOSE_EVENT = 2,
            CTRL_LOGOFF_EVENT = 5,
            CTRL_SHUTDOWN_EVENT = 6
        }

        private const int SW_HIDE = 0; 
        private const uint SPI_GETDESKWALLPAPER = 0x73;
        private const uint SPI_SETDESKWALLPAPER = 20;
        private const uint SPIF_UPDATEINIFILE = 0x01;
        private const uint SPIF_SENDWININICHANGE = 0x02; 
        private const int MAX_PATH = 260;

        static void Main(string[] args)
        {
            args = args.Select(x => x.ToLowerInvariant()).ToArray();

            if (args.Length < 1)
            {
                WriteUsage();
                return;
            }

            var command = args[0];
            if (command == "get")
            {
                var wallpaper = new string('\0', MAX_PATH);
                SystemParametersInfo(SPI_GETDESKWALLPAPER, MAX_PATH, wallpaper, 0);
                Console.WriteLine(wallpaper.Substring(0, wallpaper.IndexOf('\0')));
                return;
            }

            if (command == "set")
            {
                if (args.Length < 2)
                {
                    WriteUsage();
                    return;
                }

                var path = args[1];
                int interval;
                if (args.Length >= 3 && int.TryParse(args[2], out interval))
                {
                    CloseOtherInstances();

                    Timer.Interval = interval * 1000;
                    Timer.Enabled = true;
                    Timer.Tick += (sender, e) => SetRandomWallPaper(path);
                }
                else
                {
                    SetRandomWallPaper(path);
                    return;
                }
            }

            SetIcon();
            ShowWindow(ThisConsole, SW_HIDE);

            var handler = new EventHandler(Handler);
            SetConsoleCtrlHandler(handler, true);
            GC.KeepAlive(handler);

            Application.Run();
        }

        private static void Quit()
        {
            CleanUp();
            Application.Exit();
            Environment.Exit(1);
        }

        private static void CleanUp()
        {
            TrayIcon.Visible = false;
            TrayIcon.Dispose();
            Timer.Dispose();
        }

        private static bool Handler(CtrlType signal)
        {
            switch (signal)
            {
                case CtrlType.CTRL_C_EVENT:
                case CtrlType.CTRL_BREAK_EVENT:
                case CtrlType.CTRL_LOGOFF_EVENT:
                case CtrlType.CTRL_SHUTDOWN_EVENT:
                case CtrlType.CTRL_CLOSE_EVENT:
                    CleanUp();
                    return true;
            }
            return false;
        }

        private static void CloseOtherInstances()
        {
            foreach (var process in Process.GetProcessesByName(Path.GetFileNameWithoutExtension(Assembly.GetEntryAssembly().Location)))
            {
                if (process.Id == Process.GetCurrentProcess().Id)
                    continue;

                process.Kill();
                process.Dispose();
            }
        }

        private static void SetIcon()
        {
            using (var icon = Assembly.GetExecutingAssembly().GetManifestResourceStream("bground.tray.ico"))
            {
                if (icon == null)
                {
                    throw new InvalidOperationException("Icon not found");
                }

                TrayIcon.Icon = new Icon(icon);
                TrayIcon.Visible = true;
                TrayIcon.Text = "Right click to exit";
                TrayIcon.ContextMenu = ContextMenu;
                TrayIcon.ContextMenu.MenuItems.Add(0, new MenuItem("Exit", (sender, e) => Quit()));

                TrayIcon.Click += (sender, e) => typeof (NotifyIcon)
                    .GetMethod("ShowContextMenu", BindingFlags.Instance | BindingFlags.NonPublic)
                    .Invoke(TrayIcon, null);
            }
        }

        private static void WriteUsage()
        {
            Console.WriteLine("Usage: bround.exe (get|set <directory> [interval in seconds])");
        }

        private static void SetRandomWallPaper(string path)
        {
            if (!Directory.Exists(path))
            {
                Console.WriteLine("Path was not found");
                Quit();
            }

            SetWallPaperStyle();
            SetWallPaper(RandomElement(
                Directory.EnumerateFiles(path, "*.jpg"),
                new Random()));
        }

        private static void SetWallPaperStyle()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop", true))
            {
                if (key == null)
                {
                    throw new InvalidOperationException("Registry key for desktop image settings was not found.");
                }

                key.SetValue("WallpaperStyle", "6");
                key.SetValue("WallpaperTile", "0");
                key.Close();
            }
        }

        private static void SetWallPaper(string image)
        {
            if (SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, image, SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE))
                return;
            
            throw new Win32Exception();
        }

        public static T RandomElement<T>(IEnumerable<T> source, Random rng)
        {
            T current = default(T);
            int count = 0;
            foreach (T element in source)
            {
                count++;
                if (rng.Next(count) == 0)
                {
                    current = element;
                }
            }
            if (count == 0)
            {
                throw new InvalidOperationException("Sequence was empty");
            }
            return current;
        }
    }
}
