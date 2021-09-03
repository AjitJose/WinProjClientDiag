using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace WinProjClientDiag
{


    public partial class Form1 : Form
    {


        [DllImport("advapi32.dll", EntryPoint = "RegOpenKey")]
        public static extern
  int RegOpenKeyA(int hKey, string lpSubKey, ref int phkResult);

        [DllImport("advapi32.dll")]
        public static extern
           int RegCloseKey(int hKey);

        [DllImport("advapi32.dll", EntryPoint = "RegQueryInfoKey")]
        public static extern
           int RegQueryInfoKeyA(int hKey, string lpClass,
           ref int lpcbClass, int lpReserved,
           ref int lpcSubKeys, ref int lpcbMaxSubKeyLen,
           ref int lpcbMaxClassLen, ref int lpcValues,
           ref int lpcbMaxValueNameLen, ref int lpcbMaxValueLen,
           ref int lpcbSecurityDescriptor,
           ref FILETIME lpftLastWriteTime);

        [DllImport("advapi32.dll", EntryPoint = "RegEnumValue")]
        public static extern
           int RegEnumValueA(int hKey, int dwIndex,
           ref byte lpValueName, ref int lpcbValueName,
           int lpReserved, ref int lpType, ref byte lpData, ref int lpcbData);

        [DllImport("advapi32.dll", EntryPoint = "RegEnumKeyEx")]
        public static extern
           int RegEnumKeyExA(int hKey, int dwIndex,
           ref byte lpName, ref int lpcbName, int lpReserved,
           string lpClass, ref int lpcbClass, ref FILETIME lpftLastWriteTime);

        [DllImport("advapi32.dll", EntryPoint = "RegSetValueEx")]
        public static extern
           int RegSetValueExA(int hKey, string lpSubKey,
           int reserved, int dwType, ref byte lpData, int cbData);
        SortedList<string, string> installCol = new SortedList<string, string>();
        string regvalfull = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void FindInstalls(RegistryKey regKey, List<string> keys, SortedList<string, string> installed)
        {
            foreach (string key in keys)
            {
                using (RegistryKey rk = regKey.OpenSubKey(key))
                {
                    if (rk == null)
                    {
                        continue;
                    }
                    foreach (string skName in rk.GetSubKeyNames())
                    {
                        using (RegistryKey sk = rk.OpenSubKey(skName))
                        {
                            try
                            {
                                string disname = sk.GetValue("DisplayName").ToString();
                                if (disname.IndexOf("Project", 0) > 0)
                                {
                                    installed.Add(Convert.ToString(sk.GetValue("DisplayName")), Convert.ToString(sk.GetValue("InstallLocation")));
                                }
                            }
                            catch (Exception ex)
                            { }
                        }
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            listBox1.Items.Clear();
            List<string> installs = new List<string>();

            List<string> keys = new List<string>() {
  @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
  @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
};

            // The RegistryView.Registry64 forces the application to open the registry as x64 even if the application is compiled as x86 
            FindInstalls(RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64), keys, installCol);
            FindInstalls(RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64), keys, installCol);

            //  installCol.Sort(); // The list of ALL installed applications

            listBox1.Items.Clear();

            foreach (var abc in installCol)
            {

                listBox1.Items.Add(abc.Key);



            }




            foreach (var itm in listBox1.Items)
            {
                if (itm.ToString().IndexOf("Security", 0) < 0 && itm.ToString().IndexOf("SQL", 0) < 0 && itm.ToString().IndexOf("VisualStudio", 0) < 0 && itm.ToString().IndexOf("SDK", 0) < 0 && itm.ToString().IndexOf("Service", 0) < 0 && itm.ToString().IndexOf("MUI", 0) < 0 && itm.ToString().IndexOf("Update", 0) < 0)
                {
                    listBox2.Items.Add(itm);
                }
            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (var abc in installCol)
            {
                if (listBox2.SelectedItem.ToString() == abc.Key.ToString())
                {
                    try
                    {
                        string filepath = "";
                        if (listBox2.SelectedItem.ToString().IndexOf("2007", 0) > 0)
                        {
                            filepath = abc.Value + "\\Office12\\";
                        }

                        if (listBox2.SelectedItem.ToString().IndexOf("2010", 0) > 0)
                        {
                            filepath = abc.Value + "\\Office14\\";
                        }
                        if (listBox2.SelectedItem.ToString().IndexOf("2013", 0) > 0)
                        {
                            filepath = abc.Value + "\\Office15\\";
                        }

                        if (listBox2.SelectedItem.ToString().IndexOf("2016", 0) > 0)
                        {
                            filepath = abc.Value + "\\Office16\\";
                        }

                        if (listBox2.SelectedItem.ToString().IndexOf("2019", 0) > 0)
                        {
                            filepath = abc.Value + "\\Office16\\";
                        }

                        if (listBox2.SelectedItem.ToString().IndexOf("en-us", 0) > 0)
                        {
                            filepath = abc.Value + "\\root\\Office16\\";
                        }

                        string filefullpath = filepath + "WINPROJ.EXE";
                        //File.Exists
                        bool protocolhnd = File.Exists(filepath + "protocolhandler.exe");

                        var versionInfo = FileVersionInfo.GetVersionInfo(filefullpath);
                        string version = versionInfo.FileVersion;
                        textBox2.Text = version;
                        textBox1.Text = filefullpath;

                        if (protocolhnd)
                        {
                            var protversionInfo = FileVersionInfo.GetVersionInfo(filepath + "protocolhandler.exe");
                            string protversion = versionInfo.FileVersion;
                            textBox3.Text = protversion;
                        }
                        else
                        {
                            textBox3.Text = "";
                        }
                    }




                    catch (Exception ex)
                    {
                        textBox2.Text = "";
                        textBox1.Text = "";
                        textBox3.Text = "";

                    }
                    //MessageBox.Show(keyname);
                    //   listBox1.Items.Add(abc.Key);
                    //Console.WriteLine(subkey.GetValue("DisplayName"));
                }


            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Project 2013")
            {
                comboBox2.Visible = true;
                comboBox3.Visible = false;

                string loglvel = "";
                switch (comboBox2.Text)
                {
                    case "0 = LogNone":
                        loglvel = "00000000";
                        break;
                    case "1 = Assert":
                        loglvel = "00000001";
                        break;
                    case "2 = Unexpected":
                        loglvel = "00000002";
                        break;
                    case "3 = High":
                        loglvel = "00000003";
                        break;
                    case "4 = Medium":
                        loglvel = "00000004";
                        break;
                    case "5 = Verbose":
                        loglvel = "00000005";
                        break;
                    case "6 = VerboseEx":
                        loglvel = "00000006";
                        break;

                    default:
                        loglvel = "00000005";
                        break;
                }

                regvalfull = "[HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\15.0\\MS Project\\Logging] \r\n" +
"EnableLogging = dword:00000001 \r\n" +
"EnableTextFileLogging = dword:00000000 \r\n" +
"WinprojLog = " + textBox4.Text + "\r\n" +
"DebugCategory = dword:ffffffff \r\n" +
"DebugLevel =dword: " + loglvel + "\r\n" +
"DebugLoadSerCategory = dword:000003ed \r\n";

                textBox5.Text = regvalfull;
            }
            else
            {
                comboBox3.Visible = true;
                comboBox2.Visible = false;

                string loglvel = "";
                switch (comboBox3.Text)
                {
                    case "0 = None":
                        loglvel = "00000000";
                        break;
                    case "10 = Error":
                        loglvel = "00000010";
                        break;
                    case "15 = Warning":
                        loglvel = "00000015";
                        break;
                    case "50 = Info":
                        loglvel = "00000050";
                        break;
                    case "100 = Verbose":
                        loglvel = "00000100";
                        break;
                    case "200 = Spam":
                        loglvel = "00000200";
                        break;
                    default:
                        loglvel = "00000100";
                        break;
                }

                regvalfull = "[HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\16.0\\Common\\Logging] \r\n" +
"EnableLogging = dword:00000001 \r\n" +
"EnableTextFileLogging = dword:00000000 \r\n" +
"WinprojLog = " + textBox4.Text + "\r\n" +
"DebugCategory = dword:ffffffff \r\n" +
"DebugLevel = dword:" + loglvel + "\r\n" +
"DebugLoadSerCategory = dword:000003ed \r\n";

                textBox5.Text = regvalfull;



            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.Text = "";
            string loglvel = "";
            switch (comboBox3.Text)
            {
                case "0 = None":
                    loglvel = "00000000";
                    break;
                case "10 = Error":
                    loglvel = "00000010";
                    break;
                case "15 = Warning":
                    loglvel = "00000015";
                    break;
                case "50 = Info":
                    loglvel = "00000050";
                    break;
                case "100 = Verbose":
                    loglvel = "00000100";
                    break;
                case "200 = Spam":
                    loglvel = "00000200";
                    break;
                default:
                    loglvel = "00000100";
                    break;
            }

            regvalfull = "[HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\16.0\\Common\\Logging] \r\n" +
"EnableLogging = dword:00000001 \r\n" +
"EnableTextFileLogging = dword:00000000 \r\n" +
"WinprojLog = " + textBox4.Text + "\r\n" +
"DebugCategory = dword:ffffffff \r\n" +
"DebugLevel = dword:" + loglvel + "\r\n" +
"DebugLoadSerCategory = dword:000003ed \r\n";

            textBox5.Text = regvalfull;
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox5.Text = regvalfull;

            string tempDirectory = Path.Combine(Path.GetTempPath(), "WinprojLogs");
            Directory.CreateDirectory(tempDirectory);

            textBox4.Text = tempDirectory;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox3.Text = "";
            string loglvel = "";
            switch (comboBox2.Text)
            {
                case "0 = LogNone":
                    loglvel = "00000000";
                    break;
                case "1 = Assert":
                    loglvel = "00000001";
                    break;
                case "2 = Unexpected":
                    loglvel = "00000002";
                    break;
                case "3 = High":
                    loglvel = "00000003";
                    break;
                case "4 = Medium":
                    loglvel = "00000004";
                    break;
                case "5 = Verbose":
                    loglvel = "00000005";
                    break;
                case "6 = VerboseEx":
                    loglvel = "00000006";
                    break;

                default:
                    loglvel = "00000005";
                    break;
            }

            regvalfull = "[HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\15.0\\MS Project\\Logging] \r\n" +
"EnableLogging = dword:00000001 \r\n" +
"EnableTextFileLogging = dword:00000000 \r\n" +
"WinprojLog = " + textBox4.Text + "\r\n" +
"DebugCategory = dword:ffffffff \r\n" +
"DebugLevel = dword:" + loglvel + "\r\n" +
"DebugLoadSerCategory = dword:000003ed \r\n";


            textBox5.Text = regvalfull;
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "")
            {
                MessageBox.Show("Select the Project Version !");
            }
            else
            {
                if (comboBox1.Text == "Project 2013")
                {
                    string prjVersion = "2013";

                    string loglvel = "";

                    if (comboBox2.Text != "")
                    {
                        switch (comboBox2.Text)
                        {
                            case "0 = LogNone":
                                loglvel = "00000000";
                                break;
                            case "1 = Assert":
                                loglvel = "00000001";
                                break;
                            case "2 = Unexpected":
                                loglvel = "00000002";
                                break;
                            case "3 = High":
                                loglvel = "00000003";
                                break;
                            case "4 = Medium":
                                loglvel = "00000004";
                                break;
                            case "5 = Verbose":
                                loglvel = "00000005";
                                break;
                            case "6 = VerboseEx":
                                loglvel = "00000006";
                                break;

                            default:
                                loglvel = "00000005";
                                break;
                        }
                    }
                    else loglvel = "00000005";

                    List<string> logcat = new List<string>();
                    foreach (var listItem in listBox3.SelectedItems)
                    {
                        logcat.Add(listItem.ToString());
                    }

                    EnableLogging(prjVersion, loglvel, logcat);



                }
                else
                {
                    string prjVersion = "Others";


                    string loglvel = "";

                    if (comboBox3.Text != "")
                    {
                        switch (comboBox3.Text)
                        {
                            case "0 = None":
                                loglvel = "00000000";
                                break;
                            case "10 = Error":
                                loglvel = "00000010";
                                break;
                            case "15 = Warning":
                                loglvel = "00000015";
                                break;
                            case "50 = Info":
                                loglvel = "00000050";
                                break;
                            case "100 = Verbose":
                                loglvel = "00000100";
                                break;
                            case "200 = Spam":
                                loglvel = "00000200";
                                break;
                            default:
                                loglvel = "00000100";
                                break;
                        }
                    }
                    else loglvel = "00000100";

                    List<string> logcat = new List<string>();
                    foreach (var listItem in listBox3.SelectedItems)
                    {
                        logcat.Add(listItem.ToString());
                    }
                    DisableLogging("ld");
                    EnableLogging(prjVersion, loglvel, logcat);

                }
            }
        }

        private void DisableLogging(string src)
        {
            try
            {
                RegistryKey currentUser = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, RegistryView.Registry64);




                var reg = currentUser.OpenSubKey("Software\\Microsoft\\Office\\15.0\\MS Project\\Logging", true);

                if (reg != null)
                {
                    currentUser.DeleteSubKey("Software\\Microsoft\\Office\\15.0\\MS Project\\Logging");
                }





                reg = currentUser.OpenSubKey("Software\\Microsoft\\Office\\16.0\\Common\\Logging", true);
                if (reg != null)
                {


                    if (reg.GetValue("EnableLogging") != null)
                    {
                        reg.DeleteValue("EnableLogging");

                    }

                    if (reg.GetValue("DefaultMinimumSeverity") != null)
                    {
                        reg.DeleteValue("DefaultMinimumSeverity");
                    }

                    if (reg.GetValue("LogPath") != null)
                    {
                        reg.DeleteValue("LogPath");
                    }




                }
                if (src != "ld")
                {
                     MessageBox.Show("Logging has been Disabled");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


                                 
        }
            private void EnableLogging(string version, string logginglevel, List<string> loggingcategory)
        {
            try
            {

                if (version == "2013")
                {
                    const string userRoot = "HKEY_CURRENT_USER";
                    const string subkey = "\\Software\\Microsoft\\Office\\15.0\\MS Project\\Logging";


                    RegistryKey currentUser = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, RegistryView.Registry64);

                    var reg = currentUser.OpenSubKey("Software\\Microsoft\\Office\\15.0\\MS Project\\Logging", true);
                    if (reg == null)
                    {
                        reg = currentUser.CreateSubKey("Software\\Microsoft\\Office\\15.0\\MS Project\\Logging");
                    }

                    if (reg.GetValue("EnableLogging") == null)
                    {
                        reg.SetValue("EnableLogging", 1, RegistryValueKind.DWord);
                    }

                    if (reg.GetValue("EnableTextFileLogging") == null)
                    {
                        reg.SetValue("EnableTextFileLogging", 0, RegistryValueKind.DWord);
                    }

                    if (reg.GetValue("WinprojLog") == null)
                    {
                        reg.SetValue("WinprojLog", textBox4.Text, RegistryValueKind.String);
                    }

                    if (reg.GetValue("DebugCategory") == null)
                    {
                        reg.SetValue("DebugCategory", "ffffffff", RegistryValueKind.DWord);
                    }

                    if (reg.GetValue("DebugLevel") == null)
                    {
                        reg.SetValue("DebugLevel", logginglevel, RegistryValueKind.DWord);
                    }

                    if (reg.GetValue("DebugLoadSerCategory") == null)
                    {
                        reg.SetValue("DebugLoadSerCategory", "000003ed", RegistryValueKind.DWord);
                    }

                    if (loggingcategory.Count > 0)
                    {
                        foreach (var cat in loggingcategory)
                        {
                            if (reg.GetValue(cat) == null)
                            {
                                reg.SetValue(cat, "00000005", RegistryValueKind.DWord);
                            }

                        }
                    }

                }
                else
                {



                    RegistryKey currentUser = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, RegistryView.Registry64);

                    var reg = currentUser.OpenSubKey("Software\\Microsoft\\Office\\16.0\\Common\\Logging", true);
                    if (reg == null)
                    {
                        reg = currentUser.CreateSubKey("Software\\Microsoft\\Office\\16.0\\Common\\Logging");
                    }

                    if (reg.GetValue("EnableLogging") == null)
                    {
                        reg.SetValue("EnableLogging", 1, RegistryValueKind.DWord);
                    }

                    if (reg.GetValue("DefaultMinimumSeverity") == null)
                    {
                        reg.SetValue("DefaultMinimumSeverity", logginglevel, RegistryValueKind.DWord);
                    }

                    if (reg.GetValue("LogPath") == null)
                    {
                        reg.SetValue("LogPath", textBox4.Text, RegistryValueKind.String);
                    }


                    if (loggingcategory.Count > 0)
                    {
                        foreach (var cat in loggingcategory)
                        {
                            if (reg.GetValue(cat) == null)
                            {
                                reg.SetValue(cat, "00000005", RegistryValueKind.DWord);
                            }

                        }
                    }

                }
                MessageBox.Show("Logging has been enabled. Restart Project Client. Log files are created in" + textBox4.Text + ". Make sure you disable logging after collecting the logs");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DisableLogging("normal");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox7.Text != "")
                {
                    RegistryKey currentUser = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, RegistryView.Registry64);

                    var reg = currentUser.OpenSubKey("Software\\Microsoft\\Office\\16.0\\MS Project\\Settings", true);
                    if (reg == null)
                    {
                        reg = currentUser.CreateSubKey("Software\\Microsoft\\Office\\16.0\\MS Project\\Settings");
                    }

                    if (reg.GetValue("Timeout") == null)
                    {
                        int number;

                        bool isParsable = Int32.TryParse(textBox7.Text, out number);
                        if (isParsable)
                            reg.SetValue("Timeout", number, RegistryValueKind.DWord);
                    }
                    MessageBox.Show("Timeout value has been updated");
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                RegistryKey currentUser = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, RegistryView.Registry64);

                var reg = currentUser.OpenSubKey("Software\\Microsoft\\Office\\16.0\\MS Project\\Settings", true);
                if (reg == null)
                {
                    reg = currentUser.CreateSubKey("Software\\Microsoft\\Office\\16.0\\MS Project\\Settings");
                }

                if (reg.GetValue("Timeout") != null)
                {
                    int number;

                    bool isParsable = Int32.TryParse(reg.GetValue("Timeout").ToString(), out number);
                    if (isParsable)
                    { textBox6.Text = reg.GetValue("Timeout").ToString(); }
                    else textBox6.Text = "0";


                }
            }
            catch(Exception ex)
            {

            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
        (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }


        readonly string exportTo = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        readonly string registryLocation = @"HKEY_CURRENT_USER\Software\7-Zip";
        private static void ExportRegistry(string exportPath, string registryPath, string exportFileName)
        {
            Process regProcess = new Process();
            try
            {
                regProcess.StartInfo.FileName = "regedit.exe";
                regProcess.StartInfo.UseShellExecute = false;
                regProcess = Process.Start("regedit.exe", "/e " + Path.Combine(exportPath, exportFileName) + " " + registryPath + "");
                regProcess.WaitForExit();
            }
            catch (ArgumentNullException ane)
            {
                /* Handle error here */
            }
            catch (InvalidOperationException ioe)
            {
                /* Handle error here */
            }
            catch (Exception ex)
            {
                /* Handle all other errors here */
                regProcess.Dispose();
            }
            if (regProcess.HasExited ? true : false)
                regProcess.Dispose();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            string regLocation;
            Process regProcess = new Process();
            try
            {
                regProcess.StartInfo.FileName = "regedit.exe";
                regProcess.StartInfo.UseShellExecute = true;

                string exportTo = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                richTextBox1.Text = "Starting Export.... " + Environment.NewLine;
                richTextBox1.Text = richTextBox1.Text + Environment.NewLine + "Exporting: HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\xx.0\\Common\\SignIn\\";
                regLocation = @"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\SignIn";

                regProcess = Process.Start("regedit.exe", "/e " + Path.Combine(exportTo, "temp.txt") + " " + regLocation + "");
                regProcess.WaitForExit();

                // ExportRegistry(exportTo, regLocation, "temp.txt");

                string myfile = exportTo + "\\temp.txt";

                string newExport = exportTo + "\\ExportedKeys.txt";

                // Appending the given texts
                using (StreamWriter sw = File.AppendText(newExport))
                {
                    StreamReader sr = File.OpenText(myfile);
                    sw.WriteLine(sr.ReadToEnd());
                    sr.Close();
                }

                //Next key

                richTextBox1.Text = richTextBox1.Text + Environment.NewLine + "Exporting: HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Office\\15.0\\Common\\Identity";
                regLocation = @"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\15.0\Common\Identity";
                // ExportRegistry(exportTo, regLocation, "temp.txt");

                regProcess = Process.Start("regedit.exe", "/e " + Path.Combine(exportTo, "temp.txt") + " " + regLocation + "");



                myfile = exportTo + "\\temp.txt";

                newExport = exportTo + "\\ExportedKeys.txt";

                // Appending the given texts
                using (StreamWriter sw = File.AppendText(newExport))
                {
                    StreamReader sr = File.OpenText(myfile);
                    sw.WriteLine(sr.ReadToEnd());
                    sr.Close();
                }

                regProcess.WaitForExit();



                //NoDomainUser

                richTextBox1.Text = richTextBox1.Text + Environment.NewLine + "Exporting: HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Office\\16.0\\Common\\Identity";
                regLocation = @"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\Identity";
                // ExportRegistry(exportTo, regLocation, "temp.txt");

                regProcess = Process.Start("regedit.exe", "/e " + Path.Combine(exportTo, "temp.txt") + " " + regLocation + "");



                myfile = exportTo + "\\temp.txt";

                newExport = exportTo + "\\ExportedKeys.txt";

                // Appending the given texts
                using (StreamWriter sw = File.AppendText(newExport))
                {
                    StreamReader sr = File.OpenText(myfile);
                    sw.WriteLine(sr.ReadToEnd());
                    sr.Close();
                }

                regProcess.WaitForExit();


                //UseOnlineContent

                richTextBox1.Text = richTextBox1.Text + Environment.NewLine + "Exporting: HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\16.0\\Common\\Internet";
                regLocation = @"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\Internet";
                // ExportRegistry(exportTo, regLocation, "temp.txt");

                regProcess = Process.Start("regedit.exe", "/e " + Path.Combine(exportTo, "temp.txt") + " " + regLocation + "");



                myfile = exportTo + "\\temp.txt";

                newExport = exportTo + "\\ExportedKeys.txt";

                // Appending the given texts
                using (StreamWriter sw = File.AppendText(newExport))
                {
                    StreamReader sr = File.OpenText(myfile);
                    sw.WriteLine(sr.ReadToEnd());
                    sr.Close();
                }

                regProcess.WaitForExit();



                //UseOnlineContent

                richTextBox1.Text = richTextBox1.Text + Environment.NewLine + "Exporting: HKEY_CURRENT_USER\\Software\\Policies\\Microsoft\\Office\\16.0\\Common\\Internet";
                regLocation = @"HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Internet";
                // ExportRegistry(exportTo, regLocation, "temp.txt");

                regProcess = Process.Start("regedit.exe", "/e " + Path.Combine(exportTo, "temp.txt") + " " + regLocation + "");



                myfile = exportTo + "\\temp.txt";

                newExport = exportTo + "\\ExportedKeys.txt";

                // Appending the given texts
                using (StreamWriter sw = File.AppendText(newExport))
                {
                    StreamReader sr = File.OpenText(myfile);
                    sw.WriteLine(sr.ReadToEnd());
                    sr.Close();
                }

                regProcess.WaitForExit();


                //CDNBaseURL

                richTextBox1.Text = richTextBox1.Text + Environment.NewLine + "Exporting: HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration\\";
                regLocation = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration";
                // ExportRegistry(exportTo, regLocation, "temp.txt");

                regProcess = Process.Start("regedit.exe", "/e " + Path.Combine(exportTo, "temp.txt") + " " + regLocation + "");



                myfile = exportTo + "\\temp.txt";

                newExport = exportTo + "\\ExportedKeys.txt";

                // Appending the given texts
                using (StreamWriter sw = File.AppendText(newExport))
                {
                    StreamReader sr = File.OpenText(myfile);
                    sw.WriteLine(sr.ReadToEnd());
                    sr.Close();
                }

                regProcess.WaitForExit();

                //OfficeMgmtCOM

                richTextBox1.Text = richTextBox1.Text + Environment.NewLine + "Exporting: HKEY_LOCAL_MACHINE\\SOFTWARE\\Policies\\Microsoft\\office\\16.0\\common\\officeupdate";
                regLocation = @"HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate";
                // ExportRegistry(exportTo, regLocation, "temp.txt");

                regProcess = Process.Start("regedit.exe", "/e " + Path.Combine(exportTo, "temp.txt") + " " + regLocation + "");



                myfile = exportTo + "\\temp.txt";

                newExport = exportTo + "\\ExportedKeys.txt";

                // Appending the given texts
                using (StreamWriter sw = File.AppendText(newExport))
                {
                    StreamReader sr = File.OpenText(myfile);
                    sw.WriteLine(sr.ReadToEnd());
                    sr.Close();
                }

                regProcess.WaitForExit();


                //OfficeMgmtCOM

                richTextBox1.Text = richTextBox1.Text + Environment.NewLine + "Exporting: HKEY_LOCAL_MACHINE\\software\\policies\\microsoft\\office\\16.0\\common\\OfficeUpdate";
                regLocation = @"HKEY_LOCAL_MACHINE\software\policies\microsoft\office\16.0\common\OfficeUpdate";
                // ExportRegistry(exportTo, regLocation, "temp.txt");

                regProcess = Process.Start("regedit.exe", "/e " + Path.Combine(exportTo, "temp.txt") + " " + regLocation + "");



                myfile = exportTo + "\\temp.txt";

                newExport = exportTo + "\\ExportedKeys.txt";

                // Appending the given texts
                using (StreamWriter sw = File.AppendText(newExport))
                {
                    StreamReader sr = File.OpenText(myfile);
                    sw.WriteLine(sr.ReadToEnd());
                    sr.Close();
                }

                regProcess.WaitForExit();


                //MSProject 9

                richTextBox1.Text = richTextBox1.Text + Environment.NewLine + "Exporting: HKEY_CLASSES_ROOT\\ms-project";
                regLocation = @"HKEY_CLASSES_ROOT\ms-project";
                // ExportRegistry(exportTo, regLocation, "temp.txt");

                regProcess = Process.Start("regedit.exe", "/e " + Path.Combine(exportTo, "temp.txt") + " " + regLocation + "");



                myfile = exportTo + "\\temp.txt";

                newExport = exportTo + "\\ExportedKeys.txt";

                // Appending the given texts
                using (StreamWriter sw = File.AppendText(newExport))
                {
                    StreamReader sr = File.OpenText(myfile);
                    sw.WriteLine(sr.ReadToEnd());
                    sr.Close();
                }

                regProcess.WaitForExit();


                //MSProject.Project.9

                richTextBox1.Text = richTextBox1.Text + Environment.NewLine + "Exporting: HKEY_CLASSES_ROOT\\MSProject.Project.9";
                regLocation = @"HKEY_CLASSES_ROOT\MSProject.Project.9";
                // ExportRegistry(exportTo, regLocation, "temp.txt");

                regProcess = Process.Start("regedit.exe", "/e " + Path.Combine(exportTo, "temp.txt") + " " + regLocation + "");



                myfile = exportTo + "\\temp.txt";

                newExport = exportTo + "\\ExportedKeys.txt";

                // Appending the given texts
                using (StreamWriter sw = File.AppendText(newExport))
                {
                    StreamReader sr = File.OpenText(myfile);
                    sw.WriteLine(sr.ReadToEnd());
                    sr.Close();
                }

                regProcess.WaitForExit();


                richTextBox1.Text = richTextBox1.Text + Environment.NewLine + "Exporting Completed. Exported file is placed on desktop. Filename : ExportedKeys.txt";

            }

            catch (ArgumentNullException ane)
            {
                /* Handle error here */
            }
            catch (InvalidOperationException ioe)
            {
                /* Handle error here */
            }
            catch (Exception ex)
            {
                /* Handle all other errors here */
                regProcess.Dispose();
            }
            if (regProcess.HasExited ? true : false)
                regProcess.Dispose();


        }


    }
}
