using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;
using System.Management.Automation.Runspaces;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Drawing.Drawing2D;

namespace SETRIX
{
    public partial class Form1 : Form
    {

        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();

        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("User32.dll")]
        private static extern IntPtr GetWindowDC(IntPtr hWnd);

        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect,     // x-coordinate of upper-left corner
            int nTopRect,      // y-coordinate of upper-left corner
            int nRightRect,    // x-coordinate of lower-right corner
            int nBottomRect,   // y-coordinate of lower-right corner
            int nWidthEllipse, // width of ellipse
            int nHeightEllipse // height of ellipse
        );

        private Form activeForm;
        private Form activateForm;

        public Form1()
        {

            InitializeComponent();
            this.Text = string.Empty;
            this.ControlBox = false;
            this.FormBorderStyle = FormBorderStyle.None;
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 15, 15));

        }

        //protected override void WndProc(ref Message m)
        //{
        //    base.WndProc(ref m);
        //    const int WM_NCPAINT = 0x85;
        //    if (m.Msg == WM_NCPAINT)
        //    {
        //        IntPtr hdc = GetWindowDC(m.HWnd);
        //        if ((int)hdc != 0)
        //        {
        //            Graphics g = Graphics.FromHdc(hdc);
        //            g.FillRectangle(Brushes.Black, new Rectangle(20, 20, 20, 20));
        //            g.Flush();
        //            ReleaseDC(m.HWnd, hdc);
        //        }
        //    }
        //}

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

        private void OpenChildForm(Form childForm, object btnSender)
        {
            if (activeForm != null)
            {
                activeForm.Close();
            }
            ActivateButton(btnSender);
            activateForm = childForm;
            childForm.TopLevel = false;
            childForm.FormBorderStyle = FormBorderStyle.None;
            childForm.Dock = DockStyle.Fill;
            this.panelDesktopPane.Controls.Add(childForm);
            this.panelDesktopPane.Tag = childForm;
            childForm.BringToFront();
            childForm.Show();
        }

        private void ActivateButton(object btnSender)
        {
            
        }

        private void btnHome_Click(object sender, EventArgs e)
        {
            OpenChildForm(new Forms.FormHome(), sender);
            //lblTitle.Text = "Home";
        }

        private void btnHome_N_Click(object sender, EventArgs e)
        {
            OpenChildForm(new Forms.FormHome_N(), sender);
            //lblTitle.Text = "Home N";
        }

        private void btnHome_O_L_Click(object sender, EventArgs e)
        {
            OpenChildForm(new Forms.FormHome_S_L(), sender);
            lblTitle.Text = "Home Single Language";
        }

        private void btnHome_C_S_Click(object sender, EventArgs e)
        {
            OpenChildForm(new Forms.FormHome_C_S(), sender);
            lblTitle.Text = "Home Country Specific";
        }

        private void btnPro_Click(object sender, EventArgs e)
        {
            OpenChildForm(new Forms.FormPro(), sender);
            lblTitle.Text = "Professional";
        }

        private void btnPro_N_Click(object sender, EventArgs e)
        {
            OpenChildForm(new Forms.FormPro_N(), sender);
            lblTitle.Text = "Professional N";
        }

        private void btnEdu_Click(object sender, EventArgs e)
        {
            OpenChildForm(new Forms.FormEdu(), sender);
            lblTitle.Text = "Education";
        }

        private void btnEdu_N_Click(object sender, EventArgs e)
        {
            OpenChildForm(new Forms.FormEdu_N(), sender);
            lblTitle.Text = "Education N";
        }

        private void btnEp_Click(object sender, EventArgs e)
        {
            OpenChildForm(new Forms.FormEp(), sender);
            lblTitle.Text = "Enterprise";
        }

        private void btnEp_N_Click(object sender, EventArgs e)
        {
            OpenChildForm(new Forms.FormEp_N(), sender);
            lblTitle.Text = "Enterprise N";
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void lblTitle_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void labelTitle_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            timer2.Start();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Normal)
                this.WindowState = FormWindowState.Maximized;
            else
                this.WindowState = FormWindowState.Normal;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (Opacity == 1)
            {
                timer1.Stop();
            }
            Opacity += .2;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (Opacity <= 0)
            {
                this.Close();
            }
            Opacity -= .2;
        }

        private void label2_MouseHover(object sender, EventArgs e)
        {
            
        }
    }
}
