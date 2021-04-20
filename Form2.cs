using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CefSharp;
using CefSharp.WinForms;

namespace MencapART
{
    public partial class Form2 : Form
    {
        Timer Timer1 = new Timer();
        public ChromiumWebBrowser browser;

        public void InitBrowser()
        {
            Cef.Initialize(new CefSettings());
            browser = new ChromiumWebBrowser("http://mencapbot.mybluemix.net/#!/home");
            browser.Size = new Size(400, 400);
            browser.Location = new Point(200, 200);
            this.Controls.Add(browser);
            browser.Dock = DockStyle.Fill;
        }

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            InitBrowser();
            DocumentCarrier.value = 12;
            InitializeTimer();
        }

        private void InitializeTimer()
        {
            // Call this procedure when the application starts.  
            // Set to 1 second.  
            Timer1.Interval = 1000;
            Timer1.Tick += new EventHandler(Timer1_Tick);

            // Enable timer.
            Timer1.Enabled = true;
        }

        private void Timer1_Tick(object Sender, EventArgs e)
        {
            // Set the caption to the current time.  
            //Console.WriteLine("tick tock on the clock");

            if (DocumentCarrier.reloadPage)
            {
                browser.Reload(true);
                DocumentCarrier.reloadPage = false;
            }
            if (DocumentCarrier.pullPage)
            {
                //browser.GetBrowser().MainFrame.ViewSource();

                // Get the html source code from the main Frame.
                // This is displaying only code in the main frame and not any child frames of it.
                Task<String> taskHtml = browser.GetBrowser().MainFrame.GetSourceAsync();

                string response = taskHtml.Result;
                DocumentCarrier.documentValue = response;
                DocumentCarrier.pullPage = false;
            }
        }
    }
}
