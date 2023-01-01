using System;
using System.Data;
using System.Threading;
using System.Windows.Forms;
using WebSensor.Class;
using System.Data.SqlClient;
using System.Drawing;
using System.Media;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace WebSensor
{
    public partial class Form1 : Form
    {
        // =======================================
        #region ' Variables Defining '
        WebSensorProcess webSensor, webSensor1;
        Thread thread;
        SoundPlayer player;
        SqlConnection connection;
        SqlDataAdapter da;
        DataTable table;
        public bool bTolerance1, bTolerance2;
        private string whichSensor = "SensorOne";
        readonly string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\WebSensor\\Media\\alert.wav";
        #endregion
        // =======================================
        #region ' Form Methods '
        public Form1()
        {
            InitializeComponent();
        }
        // =======================================
        private void Form1_Load(object sender, EventArgs e)
        {
            connection = new SqlConnection("Data Source=10.3.25.106,1433; Initial Catalog=webSensor;Persist Security Info=False; User ID=sensorUser; Password=Mb82XaLz;");
            CheckForIllegalCrossThreadCalls = false;
            player = new SoundPlayer(path);
            comboBox1.SelectedIndex = 0;
            webSensor = new WebSensorProcess("http://192.168.1.213/values.xml", "SensorOne", 23, 3, 50, 20, "Sıcaklık Lab.");
            webSensor1 = new WebSensorProcess("http://192.168.1.214/values.xml", "SensorTwo", 20, 1, 50, 15, "Boyut Lab.");
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd/MM/yyyy HH:mm:ss";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "dd/MM/yyyy HH:mm:ss";
            thread = new Thread(DataReading);
            thread.Start();
            timer1.Start();
        }
        // =======================================
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            thread.Abort();
            player.Stop();
            webSensor.ThreadAbord();
            webSensor1.ThreadAbord();
            Application.ExitThread();

        }
        // =======================================
        private void timer1_Tick(object sender, EventArgs e)
        {
            timeLbl.Text = DateTime.Now.ToLongTimeString();
            dateLbl.Text = DateTime.Now.ToLongDateString();
        }
        #endregion
        // =======================================
        #region ' Buttons '
        private void filtreBtn_Click(object sender, EventArgs e)
        {
            DataView dv = table.DefaultView;
            dv.RowFilter = string.Format("Date > '{0}' AND Date <= '{1}'", dateTimePicker1.Value, dateTimePicker2.Value);
            dataGridView1.DataSource = dv;
        }
        // =======================================
        private void refreshListBtn_Click(object sender, EventArgs e)
        {
            DataList(whichSensor);
        }
        // =======================================
        private void button1_Click_1(object sender, EventArgs e)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sayfa1");
                worksheet.Name = "Sıcaklık Değerleri";
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cell(1, i).Value = dataGridView1.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheet.Cell(i + 2, j + 1).Value = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }
                FolderBrowserDialog directchoosedlg = new FolderBrowserDialog();
                if (directchoosedlg.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(@directchoosedlg.SelectedPath + "\\" + DateTime.Today.ToShortDateString() + ".xlsx");
                }
            }
        }
        // =======================================
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                whichSensor = "SensorOne";
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                whichSensor = "SensorTwo";
            }
            DataList(whichSensor);
        }
        #endregion
        // =======================================
        #region ' Sub-Methods '
        public void DataReading()
        {
            while (thread.ThreadState == ThreadState.Running)
            {
                try
                {
                    List<Item> datasensor1 = new List<Item>();
                    List<Item> datasensor2 = new List<Item>();
                    if (!webSensor.isConnected)
                    {
                        datasensor1 = webSensor.GetData();
                    }
                    if (!webSensor1.isConnected)
                    {
                        datasensor2 = webSensor1.GetData();
                    }
                    if (webSensor.isInTolerance())
                    {
                        humidityLbl.ForeColor = Color.Red;
                        tempLbl.ForeColor = Color.Red;
                        pressureLbl.ForeColor = Color.Red;
                        player.Play();
                        bTolerance1 = true;
                    }
                    else
                    {
                        humidityLbl.ForeColor = Color.Black;
                        tempLbl.ForeColor = Color.Black;
                        pressureLbl.ForeColor = Color.Black;
                        bTolerance1 = false;
                    }
                    if (webSensor1.isInTolerance())
                    {
                        lblhumidty2.ForeColor = Color.Red;
                        lbltemp2.ForeColor = Color.Red;
                        lblpressure2.ForeColor = Color.Red;
                        player.Play();
                        bTolerance2 = true;
                    }
                    else
                    {
                        lblhumidty2.ForeColor = Color.Black;
                        lbltemp2.ForeColor = Color.Black;
                        lblpressure2.ForeColor = Color.Black;
                        bTolerance2 = false;
                    }
                    if (datasensor1.Count > 0 && datasensor2.Count > 0)
                    {
                        tempLbl.Text = datasensor1[0].aval + datasensor1[0].unit;
                        humidityLbl.Text = datasensor1[1].aval + datasensor1[1].unit;
                        pressureLbl.Text = datasensor1[3].aval + datasensor1[3].unit;

                        lbltemp2.Text = datasensor2[0].aval + datasensor2[0].unit;
                        lblhumidty2.Text = datasensor2[1].aval + datasensor2[1].unit;
                        lblpressure2.Text = datasensor2[3].aval + datasensor2[3].unit;
                    }
                    if (!bTolerance1 && !bTolerance2)
                        player.Stop();
                }
                catch { }
            }
            Thread.Sleep(1000);
        }
        // =======================================
        private void DataList(string data)
        {
            try
            {
                connection.Open();
                da = new SqlDataAdapter("Select *From " + data + " ORDER BY Id DESC", connection);
                table = new DataTable();
                da.Fill(table);
                dataGridView1.DataSource = table;
                connection.Close();
            }
            catch
            {
                //MessageBox.Show(ex.Message);
            }

        }
        // =======================================
        #endregion
        // =======================================
    }
}
