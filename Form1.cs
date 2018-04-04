using System;
using System.IO.Ports;
using System.Windows.Forms;

namespace ComPortReader
{
    public partial class Form1 : Form
    {
        SerialPort serialPort;
        readonly string DATE_FORMAT = "dd.MM.yyyy HH:mm:ss";

        public Form1()
        {
            InitializeComponent();
            this.FormClosed += new FormClosedEventHandler(this.MainForm_FormClosed);
            scanPorts();
        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
            try
            {
                String comName = comboPorts.SelectedItem.ToString();
                serialPort = new SerialPort();
                serialPort.PortName = comName;
                serialPort.BaudRate = 9600;
                serialPort.ReadTimeout = 500;
                serialPort.WriteTimeout = 500;
                serialPort.Open();
                serialPort.DataReceived += serialPort2_DataReceived;
                buttonStart.Enabled = false;
                buttonStop.Enabled = true;
                buttonRescan.Enabled = false;
            }
            catch (Exception ex)
            {
                serialPort = null;
                string caption = "Ошибка.";
                MessageBoxButtons buttons = MessageBoxButtons.OK;
                DialogResult result;
                result = MessageBox.Show(ex.Message, caption, buttons, MessageBoxIcon.Error);
            }
        }

        private void buttonStop_Click(object sender, EventArgs e)
        {
            closePort();
            buttonStart.Enabled = true;
            buttonStop.Enabled = false;
            buttonRescan.Enabled = true;
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            closePort();
        }

        private void closePort()
        {
            if (serialPort != null)
            {
                serialPort.Close();
                serialPort = null;
            }
        }

        private void serialPort2_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            string POT = serialPort.ReadLine();
            BeginInvoke(new LineReceivedEvent(LineReceived), POT);
        }

        private delegate void LineReceivedEvent(string POT);

        private void LineReceived(string line)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"WriteLines2.csv", true))
            {
                String now = DateTime.Now.ToString(DATE_FORMAT);
                String res = now + ";" + line;
                file.Write(res);
                textBox1.Text = line;
                textBoxUpdateTime.Text = now;
            }
        }

        private void buttonRescan_Click(object sender, EventArgs e)
        {
            scanPorts();
        }

        private void scanPorts()
        {
            comboPorts.Items.Clear();
            string[] ports = SerialPort.GetPortNames();
            for (int i = 0; i < ports.Length; i++)
            {
                comboPorts.Items.Add(ports[i]);
            }
            if (ports.Length > 0)
                comboPorts.SelectedIndex = 0;
            if (ports.Length == 1 || ports.Length == 0)
                comboPorts.Enabled = false;
        }
    }
}
