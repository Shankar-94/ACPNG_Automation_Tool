using System;
using System.Media;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;

namespace SimpleSapScraper.Frontend.Views
{
    public partial class SuccessModal : Window
    {
        [DllImport("user32.dll")]
        private static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

        [StructLayout(LayoutKind.Sequential)]
        private struct FLASHWINFO
        {
            public uint cbSize;
            public IntPtr hwnd;
            public uint dwFlags;
            public uint uCount;
            public uint dwTimeout;
        }

        private const uint FLASHW_ALL = 3;
        private const uint FLASHW_TIMERNOFG = 12;

        public SuccessModal(string title, string message)
        {
            InitializeComponent();
            TitleTextBlock.Text = title;
            MessageTextBlock.Text = message;

            Loaded += SuccessModal_Loaded;
        }

        private void SuccessModal_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                SystemSounds.Asterisk.Play();
            }
            catch
            {
                // Ignore sound errors
            }

            try
            {
                if (Owner != null)
                {
                    var handle = new WindowInteropHelper(Owner).Handle;

                    var fInfo = new FLASHWINFO
                    {
                        cbSize = (uint)Marshal.SizeOf(typeof(FLASHWINFO)),
                        hwnd = handle,
                        dwFlags = FLASHW_ALL | FLASHW_TIMERNOFG,
                        uCount = uint.MaxValue,
                        dwTimeout = 0
                    };

                    FlashWindowEx(ref fInfo);
                }
            }
            catch
            {
                // Ignore flash errors
            }

            Activate();
            Topmost = true;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}