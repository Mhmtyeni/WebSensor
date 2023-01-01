using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Net.NetworkInformation;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;

namespace WebSensor.Class
{
    class WebSensorProcess
    {
        private XDocument xDoc;
        private SqlConnection connection;
        private SqlCommand cmd;
        private string tableName, Url, Name;
        public bool isConnected = false;
        private double temp, tempTolerance, humidity, humidityTolerance;
        private List<Item> dataLists { get; set; }
        Thread thread;

        public WebSensorProcess(string _Url, string _tableName, double _temp, double _tempTolerance, double _humidity, double _humidityTolerance, string _name)
        {
            this.dataLists = new List<Item>();

            this.connection = new SqlConnection("Data Source=10.3.25.106,1433; Initial Catalog=webSensor;Persist Security Info=False; User ID=sensorUser; Password=Mb82XaLz;");
            this.tableName = _tableName;
            this.Url = _Url;
            this.temp = _temp;
            this.Name = _name;
            this.tempTolerance = _tempTolerance;
            this.humidity = _humidity;
            this.humidityTolerance = _humidityTolerance;
            thread = new Thread(DataReading);
            Connect();
        }
        public void Connect()
        {
            thread.Start();
        }
        public List<Item> GetData()
        {
            return dataLists;
        }
        private void ReadData()
        {
            try
            {
                xDoc = XDocument.Load(Url);
                dataLists.Clear();
                var temp = xDoc.Descendants("ch1")
                           .Select(o => new Item
                           {
                               name = (string)o.Element("name"),
                               unit = (string)o.Element("unit"),
                               aval = (string)o.Element("aval"),
                           }).ToList();
                var humidity = xDoc.Descendants("ch2")
                         .Select(o => new Item
                         {
                             name = (string)o.Element("name"),
                             unit = (string)o.Element("unit"),
                             aval = (string)o.Element("aval"),
                         })
                         .ToList();

                var point = xDoc.Descendants("ch3")
                           .Select(o => new Item
                           {
                               name = (string)o.Element("name"),
                               unit = (string)o.Element("unit"),
                               aval = (string)o.Element("aval"),
                           })
                           .ToList();
                var pressure = xDoc.Descendants("ch4")
                           .Select(o => new Item
                           {
                               name = (string)o.Element("name"),
                               unit = (string)o.Element("unit"),
                               aval = (string)o.Element("aval"),
                           })
                           .ToList();
                dataLists.Add(temp[0]);
                dataLists.Add(humidity[0]);
                dataLists.Add(point[0]);
                dataLists.Add(pressure[0]);

            }
            catch (Exception)
            {
            }
        }

        private void DataReading()
        {
            while (thread.ThreadState == ThreadState.Running)
            {
                if (PingHost(Url.Split('/')[2]))
                {
                    if (isConnected)
                    {
                        isConnected = false;
                        CustomMessage(Name + " sensörüyle bağlantı kuruldu.", "Sıcaklık Takip Uygulaması", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    ReadData();
                    SqlInsert(tableName, dataLists[0].aval, dataLists[1].aval, dataLists[2].aval);
                }
                else
                {
                    if (isConnected == false)
                    {
                        isConnected = true;
                        CustomMessage(Name + " sensörüyle bağlantı kurulamadı.", "Sıcaklık Takip Uygulaması", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                Thread.Sleep(1000);
            }
        }

        private void SqlInsert(string tableName, string temp, string hum, string pres)
        {
            try
            {
                cmd = new SqlCommand();
                connection.Open();
                cmd.Connection = connection;
                cmd.CommandText = "insert into " + tableName + "(Temperature,Humidity,Pressure) values (" + temp + ",'" + hum + "','" + pres + "')";
                cmd.ExecuteNonQuery();
                connection.Close();
            }
            catch (Exception ex)
            {
            }
        }
        public void ThreadAbord()
        {
            thread.Abort();
        }
        public static bool PingHost(string Address)
        {
            bool pingable = false;
            Ping pinger = null;
            try
            {
                pinger = new Ping();
                PingReply reply = pinger.Send(Address);
                pingable = reply.Status == IPStatus.Success;
            }
            catch (PingException)
            {
            }
            finally
            {
                if (pinger != null)
                {
                    pinger.Dispose();
                }
            }

            return pingable;
        }
        public bool isInTolerance()
        {
            if (dataLists.Count > 0)
            {
                double temperature = double.Parse(dataLists[0].aval, System.Globalization.CultureInfo.InvariantCulture);
                double humidityVal = double.Parse(dataLists[1].aval, System.Globalization.CultureInfo.InvariantCulture);
                if (temperature > temp + tempTolerance || temperature < temp - tempTolerance)
                    return true;
                else if (humidityVal > humidity + humidityTolerance || humidityVal < humidity - humidityTolerance)
                    return true;
            }
            return false;

        }
        private void CustomMessage(string message, string title, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            DialogResult dialog = new DialogResult();
            dialog = MessageBox.Show(message, title, buttons, icon);

        }
    }
}
public struct Item
{
    public string name;
    public string unit;
    public string aval;
}
