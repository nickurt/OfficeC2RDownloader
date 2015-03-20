using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using System.Linq;

using MessageBox = System.Windows.MessageBox;
using System.Runtime.InteropServices;
using System;
using System.Text;
using System.ComponentModel;
using System.Collections.Generic;
using System.Net;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Globalization;

namespace OfficeC2RDownloader
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public readonly string[] languages = 
        {
            "ar-sa",
            "cs-cz",
            "da-dk",
            "de-de",
            "el-gr",
            "en-us",
            "es-es",
            "et-ee",
            "fi-fi",
            "fr-fr",
            "he-il",
            "hi-in",
            "hu-hu",
            "it-it",
            "ja-jp",
            "ko-kr",
            "lt-lt",
            "lv-lv",
            "nb-no",
            "nl-nl",
            "pl-pl",
            "pt-br",
            "pt-pt",
            "ro-ro",
            "ru-ru",
            "sv-se",
            "th-th",
            "tr-tr",
            "uk-ua",
            "vi-vn",
            "zh-cn",
            "zh-tw",
        };

        public readonly string[] x86Urls =
        {
            "{0}/Office/Data/v32.cab",
            "{0}/Office/Data/v32_{1}.cab",
            "{0}/Office/Data/{1}/i320.cab",
            "{0}/Office/Data/{1}/i640.cab",
            "{0}/Office/Data/{1}/i32{3}.cab",
            "{0}/Office/Data/{1}/i64{3}.cab",
            "{0}/Office/Data/{1}/s320.cab",
            "{0}/Office/Data/{1}/s32{3}.cab",
            "{0}/Office/Data/{1}/stream.x86.{2}.dat",
            "{0}/Office/Data/{1}/stream.x86.x-none.dat",
            "{0}/Office/Data/{1}/ui32{3}.cab",
            "{0}/Office/Data/{1}/ui64{3}.cab",
            "{0}/Office/Data/{1}/workaround.cab"
        };

        public readonly string[] x64Urls =
        {
            "{0}/Office/Data/v64.cab",
            "{0}/Office/Data/v64_{1}.cab",
            "{0}/Office/Data/{1}/i640.cab",
            "{0}/Office/Data/{1}/i64{3}.cab",
            "{0}/Office/Data/{1}/s640.cab",
            "{0}/Office/Data/{1}/s64{3}.cab",
            "{0}/Office/Data/{1}/stream.x64.{2}.dat",
            "{0}/Office/Data/{1}/stream.x64.x-none.dat",
            "{0}/Office/Data/{1}/ui64{3}.cab",
            "{0}/Office/Data/{1}/workaround.cab"
        };

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // up the connection limit
            ServicePointManager.DefaultConnectionLimit = 20;

            ObservableCollection<CheckListItem> Architectures = new ObservableCollection<CheckListItem>();
            Architectures.Add(new CheckListItem() { Value = "x86", IsChecked = false });
            Architectures.Add(new CheckListItem() { Value = "x64", IsChecked = false });
            archChkLst.ItemsSource = Architectures;

            ObservableCollection<CheckListItem> Languages = new ObservableCollection<CheckListItem>();
            for (int i = 0; i < languages.Length; i++)
            {
                Languages.Add(new CheckListItem() { Value = languages[i], IsChecked = false });
            }
            langChkLst.ItemsSource = Languages;
        }

        private void setupBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.FileName = "Setup";
            ofd.DefaultExt = ".exe";
            ofd.Filter = "Setup file|setup.exe";
            ofd.Multiselect = false;
            ofd.CheckFileExists = true;

            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                setupTxt.Text = ofd.FileName;
            }
        }

        private void outputBtn_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.ShowNewFolderButton = true;

            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                outputTxt.Text = fbd.SelectedPath;
            }
        }

        private void downloadBtn_Click(object sender, RoutedEventArgs e)
        {
            string setupPath = setupTxt.Text;
            if (!File.Exists(setupPath))
            {
                MessageBox.Show("Please enter a valid path for the setup.exe", "Setup.exe not valid", MessageBoxButton.OK, MessageBoxImage.Error);
                setupTxt.Focus();
                return;
            }

            int major = 0, minor = 0, build = 0, revision = 0;
            if (!int.TryParse(majorTxt.Text, out major))
            {
                MessageBox.Show("Please a valid major version number", "Version number invalid", MessageBoxButton.OK, MessageBoxImage.Error);
                majorTxt.Focus();
                return;
            }
            if (!int.TryParse(minorTxt.Text, out minor))
            {
                MessageBox.Show("Please a valid minor version number", "Version number invalid", MessageBoxButton.OK, MessageBoxImage.Error);
                minorTxt.Focus();
                return;
            }
            if (!int.TryParse(buildTxt.Text, out build))
            {
                MessageBox.Show("Please a valid build number", "Version number invalid", MessageBoxButton.OK, MessageBoxImage.Error);
                buildTxt.Focus();
                return;
            }
            if (!int.TryParse(revisionTxt.Text, out revision))
            {
                MessageBox.Show("Please a valid revision number", "Version number invalid", MessageBoxButton.OK, MessageBoxImage.Error);
                revisionTxt.Focus();
                return;
            }

            string[] arch = archChkLst.ItemsSource.Cast<CheckListItem>().Where(c => c.IsChecked).Select(c => c.Value).ToArray();
            if (arch.Length == 0)
            {
                MessageBox.Show("Please select at least one architecture", "No architecture selected", MessageBoxButton.OK, MessageBoxImage.Error);
                archChkLst.Focus();
                return;
            }

            string[] lang = langChkLst.ItemsSource.Cast<CheckListItem>().Where(c => c.IsChecked).Select(c => c.Value).ToArray();
            if (lang.Length == 0)
            {
                MessageBox.Show("Please select at least one language", "No language selected", MessageBoxButton.OK, MessageBoxImage.Error);
                langChkLst.Focus();
                return;
            }

            string outputPath = outputTxt.Text;
            if (!Directory.Exists(outputPath))
            {
                MessageBox.Show("Please enter a valid path for the output folder", "Output path not valid", MessageBoxButton.OK, MessageBoxImage.Error);
                outputTxt.Focus();
                return;
            }

            string baseUrl = GetUrlFromExe(setupPath);

            List<DownloadItem> Downloads = new List<DownloadItem>();

            foreach (var a in arch)
            {
                foreach (var l in lang)
                {
                    Downloads.Add(new DownloadItem()
                    {
                        Architecture = a,
                        Language = l,
                        Version = string.Format("{0}.{1}.{2}.{3}", major, minor, build, revision)
                    });
                }
            }

            foreach (DownloadItem di in Downloads)
            {
                string[] urls = di.Architecture == "x86" ? x86Urls : x64Urls;

                string buildName = string.Format("{0}_{1}_{2}", di.Version, di.Architecture, di.Language);
                string folderRoot = Path.Combine(outputPath, buildName);
                string folderOffice = Path.Combine(folderRoot, "Office");
                string folderData = Path.Combine(folderOffice, "Data");
                string folderVersion = Path.Combine(folderData, di.Version);

                Directory.CreateDirectory(folderRoot);
                Directory.CreateDirectory(folderOffice);
                Directory.CreateDirectory(folderData);
                Directory.CreateDirectory(folderVersion);

                Task t = Task.Run(new Action(() =>
                {
                    for (int i = 0; i < urls.Length; i++)
                    {
                        string downloadUrl = string.Format(urls[i], baseUrl, di.Version, di.Language, LcidFromCulture(di.Language));
                        string localPath = Path.Combine(i < 2 ? folderData : folderVersion, downloadUrl.Substring(downloadUrl.LastIndexOf("/") + 1));

                        HttpWebRequest hwreq = HttpWebRequest.CreateHttp(downloadUrl);
                        try
                        {
                            using (HttpWebResponse hwres = hwreq.GetResponse() as HttpWebResponse)
                            using (Stream hwStr = hwres.GetResponseStream())
                            using (FileStream fStr = new FileStream(localPath, FileMode.Create))
                            {
                                hwStr.CopyTo(fStr);
                            }
                        }
                        catch (WebException wex)
                        {
                            if (wex.Response == null)
                            {
                                continue;
                            }
                            if ((wex.Response as HttpWebResponse).StatusCode == HttpStatusCode.NotFound)
                            {
                                continue;
                            }
                            else
                            {
                                MessageBox.Show(wex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                        }
                    }
                }));

                t.ContinueWith(new Action<Task>((Task task) =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        MessageBox.Show(buildName + " Download complete", "Download complete", MessageBoxButton.OK, MessageBoxImage.Information);
                    });
                }));

                File.Copy(setupPath, Path.Combine(folderRoot, "setup.exe"));

                XDocument xmlDoc = new XDocument();
                XElement xmlConfig = new XElement(XName.Get("Configuration"));
                XElement xmlAdd = new XElement(XName.Get("Add"));
                xmlAdd.SetAttributeValue(XName.Get("OfficeClientEdition"), di.Architecture == "x86" ? 32 : 64);
                XElement xmlProduct = new XElement(XName.Get("Product"));
                xmlProduct.SetAttributeValue(XName.Get("ID"), "MondoRetail");
                XElement xmlLanguage = new XElement(XName.Get("Language"));
                xmlLanguage.SetAttributeValue(XName.Get("ID"), di.Language);

                xmlProduct.Add(xmlLanguage);
                xmlAdd.Add(xmlProduct);
                xmlConfig.Add(xmlAdd);
                xmlDoc.Add(xmlConfig);
                xmlDoc.Save(Path.Combine(folderRoot, "configuration.xml"));
            }
        }

        private string GetUrlFromExe(string exePath)
        {
            IntPtr libPtr = LoadLibrary(exePath);

            if (libPtr == IntPtr.Zero)
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }

            StringBuilder sb = new StringBuilder(255);

            LoadString(libPtr, 3000, sb, sb.Capacity + 1);

            FreeLibrary(libPtr);

            return sb.ToString();
        }

        private string LcidFromCulture(string culture)
        {
            CultureInfo ci = new CultureInfo(culture);
            return ci.LCID.ToString();
        }

        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Ansi)]
        static extern IntPtr LoadLibrary([MarshalAs(UnmanagedType.LPStr)] string lpFileName);

        [DllImport("user32.dll", CharSet = CharSet.Ansi)]
        static extern int LoadString(IntPtr hInstance, uint uID, StringBuilder lpBuffer, int nBufferMax);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool FreeLibrary(IntPtr hModule);
    }

    internal class CheckListItem
    {
        public string Value { get; set; }
        public bool IsChecked { get; set; }
    }

    internal class DownloadItem
    {
        public string Version { get; set; }
        public string Architecture { get; set; }
        public string Language { get; set; }
    }
}
