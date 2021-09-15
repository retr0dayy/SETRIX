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
    public partial class FormHome_S_L : Form
    {
        public FormHome_S_L()
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
            RunScript("slmgr /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH");
            RunScript("slmgr /skms kms8.msguides.com");
            RunScript("slmgr /ato");
        }
    }
}
