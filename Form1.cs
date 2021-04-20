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
using Excel=Microsoft.Office.Interop.Excel;
using System.Xml.Linq;
using System.IO;

namespace MencapART
{
    public partial class Form1 : Form
    {
        private KeyHandler ghk;
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;

        public void InitExcel(string filePath)
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(filePath);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explicit cast is not required here
        }

        public void WriteExcelRow(string questionDat, string responseDat, string cor, string rel, string qua, string tone, string otherComments, string corComments, string relComments, string quaComments, string toneComments)
        {
            int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            lastRow += 1;
            
            MySheet.Cells[lastRow, 1] = questionDat;
            MySheet.Cells[lastRow, 2] = responseDat;
            MySheet.Cells[lastRow, 3] = cor;
            MySheet.Cells[lastRow, 4] = corComments;
            MySheet.Cells[lastRow, 5] = rel;
            MySheet.Cells[lastRow, 6] = relComments;
            MySheet.Cells[lastRow, 7] = qua;
            MySheet.Cells[lastRow, 8] = quaComments;
            MySheet.Cells[lastRow, 9] = tone;
            MySheet.Cells[lastRow, 10] = toneComments;
            MySheet.Cells[lastRow, 11] = otherComments;
            FormatCells(lastRow);
        }

        public void FormatCells(int lastRow)
        {
            int Colour = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            if (lastRow % 2 == 0)
            {
                Colour = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            }
            MySheet.Cells[lastRow, 1].Interior.Color = Colour;
            MySheet.Cells[lastRow, 2].Interior.Color = Colour;
            MySheet.Cells[lastRow, 3].Interior.Color = Colour;
            MySheet.Cells[lastRow, 4].Interior.Color = Colour;
            MySheet.Cells[lastRow, 5].Interior.Color = Colour;
            MySheet.Cells[lastRow, 6].Interior.Color = Colour;
            MySheet.Cells[lastRow, 7].Interior.Color = Colour;
            MySheet.Cells[lastRow, 8].Interior.Color = Colour;
            MySheet.Cells[lastRow, 9].Interior.Color = Colour;
            MySheet.Cells[lastRow, 10].Interior.Color = Colour;
            MySheet.Cells[lastRow, 11].Interior.Color = Colour;
        }

        public string GAB(bool radioBut1, bool radioBut2, bool radioBut3)
        {
            string q = "";
            if (radioBut1)
            {
                q = "G";
            }
            if (radioBut2)
            {
                q = "A";
            }
            if (radioBut3)
            {
                q = "B";
            }
            return q;
        }

        public string YN(bool radioBut1, bool radioBut2)
        {
            string q = "";
            if (radioBut1)
            {
                q = "Y";
            }
            if (radioBut2)
            {
                q = "N";
            }
            return q;
        }

        public static class Constants
        {
            //windows message id for hotkey
            public const int WM_HOTKEY_MSG_ID = 0x0312;
        }

        public class KeyHandler
        {
            [DllImport("user32.dll")]
            private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

            [DllImport("user32.dll")]
            private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

            private int key;
            private IntPtr hWnd;
            private int id;

            public KeyHandler(Keys key, Form form)
            {
                this.key = (int)key;
                this.hWnd = form.Handle;
                id = this.GetHashCode();
            }

            public override int GetHashCode()
            {
                return key ^ hWnd.ToInt32();
            }

            public bool Register()
            {
                return RegisterHotKey(hWnd, id, 0, key);
            }

            public bool Unregiser()
            {
                return UnregisterHotKey(hWnd, id);
            }
        }

        bool[] radioQ1 = new bool[] { false, false, false };
        bool[] radioQ2 = new bool[] { false, false, false };
        bool[] radioQ3 = new bool[] { false, false, false };
        bool[] radioQ4 = new bool[] { false, false };
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ghk = new KeyHandler(Keys.OemQuotes, this);
            ghk.Register();
            this.WindowState = FormWindowState.Minimized;

            //THIS SECTION IS MANUAL CONFIGURATION LOADING
            //XElement file = XElement.Load("FileConfig.xml");
            //string filePath = (string)file.Element("SpreadsheetFile");
            //InitExcel(filePath);
            //END MANUAL CONFIGURATION LOADING

            //BEGIN AUTOMATIC FILE DETECTION
            string filePath = Directory.GetCurrentDirectory() + "/TestingData.xlsx";
            InitExcel(filePath);
            //END AUTOMATIC FILE DETECTION
        }

        private void HandleHotkey()
        {
            Console.WriteLine("WingoWango");
            this.WindowState = FormWindowState.Normal;
            int screenX = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Width;
            int screenY = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Height;
            int x = Cursor.Position.X;
            int y = Cursor.Position.Y;
            int posX = x;
            int posY = y;
            if (x > screenX - 450)
            {
                posX = screenX - 450;
            }
            if (y > screenY - 450)
            {
                posY = screenY - 450;
            }

            this.Location = new Point(posX, posY);
            DocumentCarrier.pullPage = true;
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == Constants.WM_HOTKEY_MSG_ID)
                HandleHotkey();
            base.WndProc(ref m);
        }

        public System.Windows.Forms.HtmlDocument GetHtmlDocument(string html)
        {
            WebBrowser browser = new WebBrowser();
            browser.ScriptErrorsSuppressed = true;
            browser.DocumentText = html;
            browser.Document.OpenNew(true);
            browser.Document.Write(html);
            browser.Refresh();
            return browser.Document;
        }

        private void SubmitBut_Click_1(object sender, EventArgs e)
        {
            Console.WriteLine("AYYY SUBMITTEED");
            Console.WriteLine(radioButton1.Checked);
            this.WindowState = FormWindowState.Minimized;
            string docVal = DocumentCarrier.documentValue;
            HtmlDocument document = GetHtmlDocument(docVal);
            //<p ng-show="text.isUser" ng-bind-html="$sce.trustAsHtml(text.text)" class="ng-binding" aria-hidden="false">hello lord</p>
            //document.GetElementById();
            //HtmlElementCollection elements = document.GetElementsByTagName("<p ng-show=\"text.isUser\" ng-bind-html=\"$sce.trustAsHtml(text.text)\" class=\"ng - binding\" aria-hidden=\"false\">");
            //Console.WriteLine(docVal);
            //int pos = docVal.IndexOf("hello lord");
            //Console.WriteLine(pos);
            //string b = docVal.Substring(pos - 110, 121);
            //Console.WriteLine(b);
            //foreach (HtmlElement elem in elements)
            //{
            //    Console.WriteLine(elem);
            //}           
            int elemPos = docVal.IndexOf("<p ng-show=\"text.isUser\" ng-bind-html=\"$sce.trustAsHtml(text.text)\" class=\"ng-binding\" aria-hidden=\"false\">");
            Console.WriteLine(elemPos);
            //int pos2 = docVal.IndexOf("how can");
            //Console.WriteLine(pos2);
            //string b2 = docVal.Substring(pos2 - 150, 161);
            //Console.WriteLine(b2);
            int watsPos = docVal.IndexOf("<p ng-show=\"text.isUser\" ng-bind-html=\"$sce.trustAsHtml(text.text)\" class=\"ng-binding ng-hide\" aria-hidden=\"true\">");
            Console.WriteLine(watsPos);

            Queue<ChatMessage> messageLog = stripMessages(docVal);
            foreach (ChatMessage msg in messageLog)
            {
                msg.display();
            }
            ChatMessage[] messageArray = messageLog.ToArray();


            DocumentCarrier.reloadPage = true;
            WriteExcelRow(
                messageArray[1].message,
                messageArray[2].message,
                GAB(radioButton1.Checked, radioButton2.Checked, radioButton5.Checked),
                GAB(radioButton3.Checked, radioButton4.Checked, radioButton6.Checked),
                GAB(radioButton7.Checked, radioButton8.Checked, radioButton9.Checked),
                YN(radioButton10.Checked, radioButton11.Checked),
                textBox1.Text,
                textBox2.Text,
                textBox3.Text,
                textBox4.Text,
                textBox5.Text
                );
        }

        public Queue<ChatMessage> stripMessages(string docData)
        {
            string watsonKey = "<p ng-show=\"text.isUser\" ng-bind-html=\"$sce.trustAsHtml(text.text)\" class=\"ng-binding ng-hide\" aria-hidden=\"true\">";
            string userKey = "<p ng-show=\"text.isUser\" ng-bind-html=\"$sce.trustAsHtml(text.text)\" class=\"ng-binding\" aria-hidden=\"false\">";
            Queue<ChatMessage> messageQueue = new Queue<ChatMessage>();
            int readThroughPos = 0;
            bool finishScrape = false;
            int posWat;
            int posUse;

            while (!finishScrape)
            {
                posUse = docData.IndexOf(userKey, readThroughPos);
                posWat = docData.IndexOf(watsonKey, readThroughPos);

                if (posUse == -1 && posWat == -1)
                {//we have scraped all the data out of the files
                    finishScrape = true;
                }
                else
                {
                    if (posUse == -1)
                    {//if posWatson is the only valid data
                        messageQueue.Enqueue(new ChatMessage(true, stripTill(docData, posWat + watsonKey.Length), posWat + watsonKey.Length));
                        readThroughPos = posWat + watsonKey.Length;
                    }
                    else
                    {
                        if (posWat == -1)
                        {//if posUser is the only valid data
                            messageQueue.Enqueue(new ChatMessage(false, stripTill(docData, posUse + userKey.Length), posUse + userKey.Length));
                            readThroughPos = posUse + userKey.Length;
                        }
                        else
                        {//in the case that both are valid data
                            if (posWat < posUse)
                            {
                                messageQueue.Enqueue(new ChatMessage(true, stripTill(docData, posWat + watsonKey.Length), posWat + watsonKey.Length));
                                readThroughPos = posWat + watsonKey.Length;
                            }
                            else
                            {
                                messageQueue.Enqueue(new ChatMessage(false, stripTill(docData, posUse + userKey.Length), posUse + userKey.Length));
                                readThroughPos = posUse + userKey.Length;
                            }
                        }
                    }
                }
            }
            return messageQueue;
        }

        public string stripTill(string doc, int startPos)
        {
            bool stopLoop = false;
            string combination = "";
            while (!stopLoop)
            {
                string c = doc.Substring(startPos, 1);
                if (c == "<")
                {
                    stopLoop = true;
                }
                else
                {
                    combination += c;
                    startPos++;
                }
            }
            return combination;
        }

        public class ChatMessage
        {
            public bool fromWatson;
            public string message;
            public int position;

            public ChatMessage(bool isWat, string mess, int index)
            {
                fromWatson = isWat;
                message = mess;
                position = index;
            }

            public void display()
            {
                string disp = fromWatson.ToString() + "_" + message + "_" + position.ToString();
                Console.WriteLine(disp);
            }
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            MyBook.Save();
            MyApp.Workbooks.Close();
            MyApp.Quit();
            Console.WriteLine("Wingo Wango Tango");
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }
    }
}
