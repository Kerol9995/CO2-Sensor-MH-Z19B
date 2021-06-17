using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        byte co2read = 0x86;
        byte getRange = 0x9B;
        byte getABC = 0x7D;
        byte CO2calibration = 0x87;
        byte SelfCalibration = 0x79;
        byte OnSelfCalibration = 0xA0;
        byte OffSelfCalibration = 0x00;
        bool newReceived = false;
        byte[] receivedData = new byte[9];
        string selectedPort;
        double total = 0;
        double recData = 0;

        private byte getCheckSum(byte[] packet)
        {
            byte checksum = 0;
            for (int i = 1; i < 8; i++)
            {
                checksum += packet[i];
            }
            checksum = (byte)(packet[0] - checksum);
            checksum += 1;
            return checksum;
        }

        private void request(byte cmd, byte value)
        {
            byte[] arr = { 0xFF, 0x01, cmd, value, 0x00, 0x00, 0x00, 0x00, 0x00 };
            arr[8] = getCheckSum(arr);
            serialPort1.Write(arr, 0, 9);
        }

        private void getports()
        {
            string[] ports = System.IO.Ports.SerialPort.GetPortNames();
            foreach (string port in ports)
            {
                comboBox1.Items.Add(port);
                //Console.WriteLine(port);
            }
        }

        public void drawData()
        {
            //Console.WriteLine(getCheckSum(receivedData) + " " + receivedData[8]);
            if (getCheckSum(receivedData) == receivedData[8])
            {
                recData++;
                int CO2 = receivedData[2] * 256 + receivedData[3];
                int temp = receivedData[4] - 40 + (int)numericUpDown2.Value;
                label1.Text = CO2.ToString() + " ppm";
                label2.Text = temp.ToString() + "°C";
                label4.Text = ((total - recData) / total).ToString("p2");
                if (CO2 < 1200)
                {
                    chart1.Series[0].Color = Color.Red;
                    label1.ForeColor = Color.Red;
                    //System.Media.SystemSounds.Exclamation.Play();
                }
                if (CO2 < 1000)
                {
                    chart1.Series[0].Color = Color.Orange;
                    label1.ForeColor = Color.Orange;
                }
                if (CO2 < 800)
                {
                    chart1.Series[0].Color = Color.LimeGreen;
                    label1.ForeColor = Color.LimeGreen;
                }
                chart1.Series[0].Points.AddXY(DateTime.Now.ToShortTimeString(), CO2);
                chart1.Series[1].Points.AddXY(DateTime.Now.ToShortTimeString(), temp);
                if (chart1.Series[0].Points.Count > 359)
                {
                    chart1.Series[0].Points.RemoveAt(0);
                    chart1.Series[1].Points.RemoveAt(0);
                }
                using (StreamWriter sw = new StreamWriter("CO2log.txt", true, System.Text.Encoding.Default))
                {
                    sw.WriteLine(DateTime.Now + "\tCO2=" + CO2 + " ppm\tT=" + temp + "°C");
                }
            }
            else
            {
                label4.Text = ((total - recData) / total).ToString("p2");
                using (StreamWriter sw = new StreamWriter("CO2log.txt", true, System.Text.Encoding.Default))
                {
                    sw.WriteLine(DateTime.Now + "\tError\t" + receivedData[0] + ',' + receivedData[1] + ',' + receivedData[2] + ',' + receivedData[3] + ',' + receivedData[4] + ',' + receivedData[5] + ',' + receivedData[6] + ',' + receivedData[7] + ',' + receivedData[8] + '\t' + getCheckSum(receivedData));
                }
                //chart1.ChartAreas[0].AxisX.ScaleView.Zoom(0, 24);
                //chart1.ChartAreas[0].AxisY.ScaleView.Zoom(400, 2000);
            }
        }

        public Form1()
        {
            InitializeComponent();
            getports();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            request(co2read, 0);
            while (!newReceived) ;
            newReceived = false;
            drawData();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (selectedPort == null) MessageBox.Show("Select COMport!");
            else
            {
                if (!serialPort1.IsOpen)
                {
                    serialPort1.PortName = selectedPort;
                    serialPort1.Open();
                    //Console.WriteLine("Open");
                    button1.Text = "Disconnect";
                    button1.BackColor = Color.Red;
                    request(co2read, 0);
                    while (!newReceived) ;
                    newReceived = false;
                    drawData();
                    timer1.Enabled = true;
                    //timer2.Enabled = true;
                    
                }
                else
                {
                    serialPort1.Close();
                    //Console.WriteLine("Close");
                    button1.Text = "Connect";
                    button1.BackColor = Color.LimeGreen;
                    newReceived = false;
                    timer1.Enabled = false;
                    //timer2.Enabled = false;
                    total = 0;
                    recData = 0;
                }
            }
        }

        private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            
            serialPort1.Read(receivedData, 0, 9);
            total++;
            newReceived = true;
            Console.WriteLine(receivedData[1] + ";" + receivedData[2] + ";" + receivedData[3] + ";" + receivedData[4] + ";" + receivedData[5] + ";" + receivedData[6] + ";" + receivedData[7] + ";" + receivedData[8]);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedPort = comboBox1.Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!serialPort1.IsOpen) MessageBox.Show("COMport Closed!");
            else
            {
                DialogResult result = MessageBox.Show("Before calibrating the zero point, please ensure that the sensor is stable for more than 20 minutes at 400ppm ambient environment.", "Warning",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2,
                MessageBoxOptions.DefaultDesktopOnly);
                if (result == DialogResult.Yes)
                {
                    if (MessageBox.Show("Are you sure?", "Question",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2,
                    MessageBoxOptions.DefaultDesktopOnly) == DialogResult.OK)
                    {
                        request(CO2calibration, 0);
                        //while (!newReceived) ;
                        //newReceived = false;
                        //if(receivedData[1] == CO2calibration & receivedData[2] == 1)
                        //{
                            MessageBox.Show("Done!!!", "Information",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information,
                                MessageBoxDefaultButton.Button1,
                                MessageBoxOptions.DefaultDesktopOnly);
                        //}
                        //else
                        //{
                        //    MessageBox.Show("No response.", "Information",
                        //        MessageBoxButtons.OK,
                        //        MessageBoxIcon.Information,
                        //        MessageBoxDefaultButton.Button1,
                        //        MessageBoxOptions.DefaultDesktopOnly);
                        //}
                    }
                }
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            timer1.Interval= (int)numericUpDown1.Value;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!serialPort1.IsOpen) MessageBox.Show("COMport Closed!");
            else
            {
                request(getABC, 0);                                    //125
                while (!newReceived) ;
                newReceived = false;
                total--;
                //Console.WriteLine(receivedData[1] + ";" + receivedData[2] + ";" + receivedData[3] + ";" + receivedData[4] + ";" + receivedData[5] + ";" + receivedData[6] + ";" + receivedData[7] + ";" + receivedData[8]);
                if (receivedData[1] == getABC & receivedData[7]==0)
                {
                    request(SelfCalibration, OnSelfCalibration);        //121
                    while (!newReceived) ;
                    newReceived = false;
                    total--;
                    if (receivedData[1]==SelfCalibration & receivedData[2] == 1)
                    {
                        button3.BackColor = Color.LimeGreen;
                        button3.Text = "On/Off Self-calibration for Zero Point (ON)";
                    }
                }
                else
                {
                    request(SelfCalibration, OffSelfCalibration);
                    while (!newReceived) ;
                    newReceived = false;
                    total--;
                    if (receivedData[1] == SelfCalibration & receivedData[2] == 1)
                    {
                        button3.BackColor = SystemColors.Control;
                        button3.Text = "On/Off Self-calibration for Zero Point (OFF)";
                    }
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (!serialPort1.IsOpen) MessageBox.Show("COMport Closed!");
            else
            {
                request(getRange, 0);
                while (!newReceived) ;
                newReceived = false;
                total--;
                button4.Text = (receivedData[4] * 256 + receivedData[5]).ToString();
                
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (!serialPort1.IsOpen) MessageBox.Show("COMport Closed!");
            else
            {
                request(getABC, 0);
                while (!newReceived) ;
                newReceived = false;
                total--;
                button5.Text = receivedData[7].ToString();
            }
        }
    }
}
