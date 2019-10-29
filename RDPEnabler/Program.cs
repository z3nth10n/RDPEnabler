using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace RDPEnabler
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            ChromeOptions options = new ChromeOptions();

            //string profilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            //    "Google", "Chrome", "User Data");
            //options.AddArgument($"user-data-dir={profilePath}");
            //options.AddArgument("disable-infobars");

            // inomeogfingihgjfjlpeplalcfajhgai
            string extensionPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Google", "Chrome", "User Data", "Default", "Extensions", "inomeogfingihgjfjlpeplalcfajhgai", "1.5_0");
            options.AddArgument("--start-maximized");
            options.AddArgument($"--load-extension={extensionPath}");

            using (IWebDriver web = new ChromeDriver(options))
            {
                web.Navigate().GoToUrl("https://remotedesktop.google.com/support");

                // TODO: Store credentials in environment or in user config (it's ignored by .gitignore)
                SendKeys.SendWait("xxx");
                SendKeys.SendWait("{ENTER}");

                Thread.Sleep(1000);

                SendKeys.SendWait("xxx");
                SendKeys.SendWait("{ENTER}");

                Thread.Sleep(2000);

                MouseOperations.SetCursorPosition(1330, 515);
                MouseOperations.Click();

                Thread.Sleep(3000);

                MouseOperations.SetCursorPosition(1010, 515);
                MouseOperations.Click();

                Thread.Sleep(1000);

                if (Clipboard.ContainsText(TextDataFormat.Text))
                {
                    string clipboardText = Clipboard.GetText(TextDataFormat.Text);
                    // Do whatever you need to do with clipboardText

                    Console.WriteLine($"Your code is: {clipboardText}");
                }

                Console.Read();
            }
        }
    }

    public class MouseOperations
    {
        [Flags]
        public enum MouseEventFlags
        {
            LeftDown = 0x00000002,
            LeftUp = 0x00000004,
            MiddleDown = 0x00000020,
            MiddleUp = 0x00000040,
            Move = 0x00000001,
            Absolute = 0x00008000,
            RightDown = 0x00000008,
            RightUp = 0x00000010
        }

        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetCursorPos(out MousePoint lpMousePoint);

        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        public static void SetCursorPosition(int x, int y)
        {
            SetCursorPos(x, y);
        }

        public static void SetCursorPosition(MousePoint point)
        {
            SetCursorPos(point.X, point.Y);
        }

        public static MousePoint GetCursorPosition()
        {
            MousePoint currentMousePoint;
            var gotPoint = GetCursorPos(out currentMousePoint);
            if (!gotPoint) { currentMousePoint = new MousePoint(0, 0); }
            return currentMousePoint;
        }

        public static void MouseEvent(MouseEventFlags value)
        {
            MousePoint position = GetCursorPosition();

            mouse_event
                ((int)value,
                    position.X,
                    position.Y,
                    0,
                    0)
                ;
        }

        public static void Click()
        {
            MouseEvent(MouseEventFlags.LeftDown);
            MouseEvent(MouseEventFlags.LeftUp);
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MousePoint
        {
            public int X;
            public int Y;

            public MousePoint(int x, int y)
            {
                X = x;
                Y = y;
            }
        }
    }
}