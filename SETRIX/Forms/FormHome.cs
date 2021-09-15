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

namespace SETRIX.Forms
{
    public partial class FormHome : Form
    {
        public FormHome()
        {
            InitializeComponent();
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

        private void button1_Click(object sender, EventArgs e)
        {
            RunScript("slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99");
            RunScript("slmgr /skms kms8.msguides.com");
            RunScript("slmgr /ato");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            RunScript("slmgr /ipk 3KHY7-WNT83-DGQKR-F7HPR-844BM");
            RunScript("slmgr /skms kms8.msguides.com");
            RunScript("slmgr /ato");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            RunScript("slmgr /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH");
            RunScript("slmgr /skms kms8.msguides.com");
            RunScript("slmgr /ato");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            RunScript("slmgr /ipk PVMJN-6DFY6-9CCP6-7BKTT-D3WVR");
            RunScript("slmgr /skms kms8.msguides.com");
            RunScript("slmgr /ato");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            RunScript("slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX");
            RunScript("slmgr /skms kms8.msguides.com");
            RunScript("slmgr /ato");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            RunScript("slmgr /ipk MH37W-N47XK-V7XM9-C7227-GCQG9");
            RunScript("slmgr /skms kms8.msguides.com");
            RunScript("slmgr /ato");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            RunScript("slmgr /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2");
            RunScript("slmgr /skms kms8.msguides.com");
            RunScript("slmgr /ato");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            RunScript("slmgr /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ");
            RunScript("slmgr /skms kms8.msguides.com");
            RunScript("slmgr /ato");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            RunScript("slmgr /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43");
            RunScript("slmgr /skms kms8.msguides.com");
            RunScript("slmgr /ato");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            RunScript("slmgr /ipk DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4");
            RunScript("slmgr /skms kms8.msguides.com");
            RunScript("slmgr /ato");
        }
    }
}
