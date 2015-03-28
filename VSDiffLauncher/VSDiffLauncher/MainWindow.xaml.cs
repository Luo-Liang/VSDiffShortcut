using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace VSDiffLauncher
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        string prevDrop;

        private void container_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // Note that you can have more than one file.
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);

                // Assuming you have one file that you care about, pass it off to whatever
                // handling code you have defined.
                if (files.Length >= 2)
                {
                    if (files.Length > 2)
                        hintBox.Text = "Launching VS. Diffing only first 2 files.";
                    else
                        hintBox.Text = "Launching VS.";
                    launchVSDiff(files);
                }
                else
                {
                    Debug.Assert(files.Length == 1);
                    if (prevDrop == null)
                    {
                        hintBox.Text = "One more file to go.";
                        prevDrop = files[0];
                    }
                    else
                    {
                        launchVSDiff(new[] { files[0], prevDrop });
                    }
                }
            }
        }

        private string grabVSPath()
        {
            try
            {
                var key = Registry.LocalMachine.OpenSubKey("SOFTWARE").OpenSubKey("Microsoft").OpenSubKey("VisualStudio");
                var version  = key.GetSubKeyNames().Where<string>(o =>
                {
                    double notUsed;
                    return double.TryParse(o, out notUsed);
                }).OrderByDescending<string, double>(o => double.Parse(o)).Take(1);
                var vsInstallPath = (String)Registry.GetValue(String.Format("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\VisualStudio\\{0}", version), "InstallDir", "");
                return string.Format("{0}devenv.exe", vsInstallPath);
            }
            catch
            {
                return null;
            }
        }

        private void launchVSDiff(string[] files)
        {
            hintBox.Text = "Launched VS.";
            try
            {
                var psi = new ProcessStartInfo(grabVSPath());
                psi.Arguments = string.Format("/diff {0} {1}", files[0], files[1]);
                psi.UseShellExecute = true;
                Process.Start(psi);
                return;
            }
            catch{     }
            if (UserSettings.Default.VSPATH != null)
            {
                try
                {
                    var psi = new ProcessStartInfo(UserSettings.Default.VSPATH);
                    psi.Arguments = string.Format("/diff {0} {1}",files[0],files[1]);
                    psi.UseShellExecute = true;
                    Process.Start(psi);
                    return;
                }
                catch
                {
                }
            }
            //Process.Start(@"C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\devenv.exe");
            hintBox.Text = "Cannot find VS. Where is VS?";
            vsPath.Visibility = System.Windows.Visibility.Visible;
        }

        private void vsPath_KeyUp(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                vsPath.Visibility = System.Windows.Visibility.Collapsed;
                UserSettings.Default.VSPATH = vsPath.Text;
                UserSettings.Default.Save();
                hintBox.Text = "Try again with your files.";
            }
        }
    }
}
