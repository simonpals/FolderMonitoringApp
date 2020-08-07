using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using Microsoft.Win32;

namespace FolderMonitoringApp
{
    public partial class Form1 : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        static List<string> searchedFiles = new List<string>();
        static int timeToUpadateMins = 1;
        static int timeToUpadateMs = timeToUpadateMins * 60 * 1000;  //  every 1 min
        static System.Windows.Forms.Timer wfTimer = new System.Windows.Forms.Timer();
        static int lastSentTimeMins = -1;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadSettings();
            UpdateFilesTable();

            wfTimer.Interval = timeToUpadateMs;
            wfTimer.Tick += new EventHandler(delegate (Object o, EventArgs a)
            {
                UpdateFilesTable();
                textBox5.Text = "";
                CheckAndSendMail();
                if (!textBox5.Text.Any())
                    textBox5.Text = "Done";
            });
            wfTimer.Start();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            wfTimer.Stop();
        }

        private void CheckAndSendMail()
        {
            bool needSend = false;
            DayOfWeek wk = DateTime.Today.DayOfWeek;
            switch (wk)
            {
                case DayOfWeek.Monday:
                    needSend = checkBox1.Checked;
                    break;
                case DayOfWeek.Tuesday:
                    needSend = checkBox2.Checked;
                    break;
                case DayOfWeek.Wednesday:
                    needSend = checkBox3.Checked;
                    break;
                case DayOfWeek.Thursday:
                    needSend = checkBox4.Checked;
                    break;
                case DayOfWeek.Friday:
                    needSend = checkBox5.Checked;
                    break;
                case DayOfWeek.Saturday:
                    needSend = checkBox6.Checked;
                    break;
                case DayOfWeek.Sunday:
                    needSend = checkBox7.Checked;
                    break;
            };

            if (needSend)
            {
                string timeNow = DateTime.Now.ToString("HH:mm");
                if (DateTime.Now.Hour == dateTimePicker1.Value.Hour &&
                    Math.Abs(DateTime.Now.Minute - dateTimePicker1.Value.Minute) <= timeToUpadateMins)
                {
                    if (Math.Abs(DateTime.Now.Minute - lastSentTimeMins) > timeToUpadateMins)
                    {
                        SendMail();
                        lastSentTimeMins = DateTime.Now.Minute;
                    }
                }
            }
        }

        bool MailFieldsCorrect()
        {
            return textBox3.Text.Any() && textBox4.Text.Any() && textBox6.Text.Any();
        }

        private void SendMail()
        {
            if (!MailFieldsCorrect()) return;

            try
            {
                MailMessage email = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");

                email.From = new MailAddress(textBox3.Text);
                email.To.Add(textBox6.Text);
                email.Subject = textBox7.Text;
                email.Body = string.Join("\n", searchedFiles);
                System.IO.File.WriteAllText("MonitoredFiles.txt", string.Join("\n", searchedFiles));
                email.Attachments.Add(new Attachment("MonitoredFiles.txt"));

                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential(textBox3.Text, textBox4.Text);
                SmtpServer.EnableSsl = true;

                SmtpServer.Send(email);
            }
            catch (Exception es)
            {
                textBox5.Text = es.Message;
                //MessageBox.Show(es.Message);
            }

        }

        private void SearchFiles()
        {
            searchedFiles.Clear();

            string path = textBox1.Text;
            string[] filters = textBox2.Text.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

            for (int index = 0; index < filters.Length; index++)
            {
                filters[index] = filters[index].Trim();
            }

            if (Directory.Exists(path))
            {
                ProcessDirectoryCountFiles(path, filters, searchedFiles);
                searchedFiles.Add("*****************************************");
                ProcessDirectory(path, filters, searchedFiles);
            }
        }

        private void UpdateFilesTable()
        {
            SearchFiles();
            listBox1.Items.Clear();

            foreach (var searchedFile in searchedFiles)
            {
                listBox1.Items.Add(searchedFile);
            }
        }

        static string fileSettingPath = "C:\\ProgramData\\settings.ini";

        private void LoadSettings()
        {
            if (!File.Exists(fileSettingPath)) return;

            using (StreamReader readtext = new StreamReader(fileSettingPath))
            {
                textBox3.Text = readtext.ReadLine();
                textBox4.Text = readtext.ReadLine();
                textBox6.Text = readtext.ReadLine();
                textBox7.Text = readtext.ReadLine();
                textBox1.Text = readtext.ReadLine();
                textBox2.Text = readtext.ReadLine();
                checkBox1.Checked = bool.Parse(readtext.ReadLine());
                checkBox2.Checked = bool.Parse(readtext.ReadLine());
                checkBox3.Checked = bool.Parse(readtext.ReadLine());
                checkBox4.Checked = bool.Parse(readtext.ReadLine());
                checkBox5.Checked = bool.Parse(readtext.ReadLine());
                checkBox6.Checked = bool.Parse(readtext.ReadLine());
                checkBox7.Checked = bool.Parse(readtext.ReadLine());
                dateTimePicker1.Value = DateTime.Parse(readtext.ReadLine());
                checkBox8.Checked = bool.Parse(readtext.ReadLine());
            }
        }

        private void SaveSettings()
        {
            using (StreamWriter writetext = new StreamWriter(fileSettingPath))
            {
                writetext.WriteLine(textBox3.Text); //sender email
                writetext.WriteLine(textBox4.Text); //sender password
                writetext.WriteLine(textBox6.Text); //receiver email
                writetext.WriteLine(textBox7.Text); //message subject
                writetext.WriteLine(textBox1.Text); //search directory
                writetext.WriteLine(textBox2.Text); //search filter            
                writetext.WriteLine(checkBox1.Checked.ToString());
                writetext.WriteLine(checkBox2.Checked.ToString());
                writetext.WriteLine(checkBox3.Checked.ToString());
                writetext.WriteLine(checkBox4.Checked.ToString());
                writetext.WriteLine(checkBox5.Checked.ToString());
                writetext.WriteLine(checkBox6.Checked.ToString());
                writetext.WriteLine(checkBox7.Checked.ToString());
                writetext.WriteLine(dateTimePicker1.Value.ToString()); // send time
                writetext.WriteLine(checkBox8.Checked.ToString()); // run app with windows
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveSettings();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog directchoosedlg = new FolderBrowserDialog();
            if (directchoosedlg.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = directchoosedlg.SelectedPath;
            }
        }

        public static void ProcessDirectoryCountFiles(string targetDirectory, string[] filters, List<string> searchFiles)
        {
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            int totalCount = fileEntries.Count();
            int filteredCount = 0;

            foreach (string fileName in fileEntries)
            {
                for (int index = 0; index < filters.Length; index++)
                {
                    if (fileName.EndsWith(filters[index]) || filters[index] == "*")
                    {
                        filteredCount++;
                    }
                }
            }
            searchFiles.Add(targetDirectory+ "  |  Total files count: "+ totalCount + "  |  Filtered files count: "+filteredCount);

            // Recurse into subdirectories of this directory. 
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessDirectoryCountFiles(subdirectory, filters, searchFiles);
        }

        public static void ProcessDirectory(string targetDirectory, string[] filters, List<string> searchFiles)
        {
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
                ProcessFile(fileName, filters, searchFiles);

            // Recurse into subdirectories of this directory. 
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessDirectory(subdirectory, filters, searchFiles);
        }

        public static void ProcessFile(string path, string[] filters, List<string> searchFiles)
        {
            for (int index = 0; index < filters.Length; index++)
            {
                if (path.EndsWith(filters[index]) || filters[index] == "*")
                {
                    searchFiles.Add(path);
                }
            }
        }


        public static void AddApplicationToStartup()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
                key.SetValue("FolderMonitoringApp", "\"" + Application.ExecutablePath + "\"");
            }
        }

        public static void RemoveApplicationFromStartup()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
                key.DeleteValue("FolderMonitoringApp", false);
            }
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {            
            if (checkBox8.Checked)
            {
                AddApplicationToStartup();
            }
            else
            {
                RemoveApplicationFromStartup();
            }
        }
    }
}
