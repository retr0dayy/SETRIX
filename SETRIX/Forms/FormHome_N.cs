using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management.Automation.Runspaces;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Threading;
using System.Net;
using System.IO;
using System.IO.Compression;
using System.Diagnostics;
using CliWrap;
using System.Reflection;

namespace SETRIX.Forms
{
    public partial class FormHome_N : Form
    {

        private ProcessStartInfo psi;
        private Process cmd;
        private delegate void InvokeWithString(string text);

        public FormHome_N()
        {
            InitializeComponent();

            this.FormClosing += Form1_FormClosing;
        }

        Process p;

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            p?.Dispose();
        }

        private string RunScript(string script)
        {
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();
            pipeline.Commands.AddScript(script);
            pipeline.Commands.Add("Out-String");
            Collection<PSObject> results = pipeline.Invoke();
            runspace.Close();
            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject pSObject in results)
                stringBuilder.AppendLine(pSObject.ToString());
            return stringBuilder.ToString();
        }



        WebClient client;

        private void button1_Click(object sender, EventArgs e)
        {
            if (p != null)
                p.Dispose();

            p = new Process();
            p.StartInfo.WorkingDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            p.StartInfo.FileName = @"Resources\OFFICE.cmd";
            p.StartInfo.Arguments = "";
            p.StartInfo.CreateNoWindow = true;
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardInput = true;
            p.OutputDataReceived += proc_OutputDataReceived;
            p.Start();
            p.BeginOutputReadLine();
        }

        private void proc_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            this.Invoke((Action)(() =>
            {
                textBox1.AppendText(Environment.NewLine + e.Data);
                textBox1.ScrollToCaret();
            }));

            //can use either of these lines.
            (sender as Process)?.StandardInput.WriteLine();
            //p.StandardInput.WriteLine();
        }

        private void FormHome_N_Load(object sender, EventArgs e)
        {
            client = new WebClient();
            client.DownloadProgressChanged += Client_DownloadProgressChanged;
            client.DownloadFileCompleted += Client_DownloadFileCompleted;
        }

        private void Client_DownloadFileCompleted(object sender, AsyncCompletedEventArgs e)
        {
            //label1.Text = "OFFICE INSTALLED SUCCESSFULLY";
            //string targetFolder = "CopyofMSOPP2k19x86x64en";
            //string sourceZipFile = "CopyofMSOPP2k19x86x64en.zip";

            //ZipFile.ExtractToDirectory(sourceZipFile, targetFolder);

            //Process.Start("CopyofMSOPP2k19x86x64en/MSOPP2k19x86x64en.exe");
        }

        private void Client_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            Invoke(new MethodInvoker(delegate ()
            {
                //progressBar1.Minimum = 0;
                double receive = double.Parse(e.BytesReceived.ToString());
                double total = double.Parse(e.TotalBytesToReceive.ToString());
                double percentage = receive / total * 100;
                label1.Text = $"Installing Office\n{string.Format("{0:0.##}", percentage)}%";
                //progressBar1.Value = int.Parse(Math.Truncate(percentage).ToString());

            }));
        }
    }
}
