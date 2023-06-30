using Microsoft.Win32;
using System.Web;
using Windows.Devices.SerialCommunication;
using Windows.Foundation;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System;
using System.Threading.Tasks;
using Windows.Devices.Enumeration;
using System.Windows.Forms;


namespace SerialName
{
    public partial class SerialName : Form
    {

        public object DeviceInformation { get; }

        public SerialName()
        {
            InitializeComponent();
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void SerialName_Load(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns[0].HeaderText = "Registry Key";
            dataGridView1.Columns[1].HeaderText = "FriendlyName";
            LoadSerialDevices("SYSTEM\\ControlSet001\\Enum\\FTDIBUS");
            LoadSerialDevices("SYSTEM\\ControlSet001\\Enum\\USB");
            LoadSerialDevices("SYSTEM\\ControlSet001\\Enum\\VSBC9");
            LoadSerialDevices("SYSTEM\\ControlSet001\\Enum\\FabulaTech");
            int dgv_width = dataGridView1.Columns.GetColumnsWidth(DataGridViewElementStates.Visible);
            this.Width = dgv_width + 120;
            this.Width = dataGridView1.Width + 50;
        }
        private void LoadSerialDevices(string key)
        {
            using RegistryKey myKey = Registry.LocalMachine.OpenSubKey(key);
            if (myKey == null) return;
            string[] subKeyNames = myKey.GetSubKeyNames();
            foreach (string s in subKeyNames)
            {
                if (s.Contains("VID_") && s.Contains("PID_") || s.Contains("VSP") || s.Contains("DEVICE"))
                {
                    RegistryKey mySubKey = myKey.OpenSubKey(s);
                    if (mySubKey == null) return;
                    string[] subSubKeyNames = mySubKey.GetSubKeyNames();
                    foreach (string ss in subSubKeyNames)
                    {
                        RegistryKey DeviceKey = mySubKey.OpenSubKey(ss);
                        if (DeviceKey == null) return;
                        string[] values = DeviceKey.GetValueNames();
                        foreach (string v in values)
                        {
                            string friendlyName = null;
                            if (DeviceKey.GetValue("FriendlyName") != null)
                            {
                                friendlyName = DeviceKey.GetValue("FriendlyName").ToString();
                                if (friendlyName == null) return;
                                if (friendlyName.Contains("COM"))
                                    dataGridView1.Rows.Add(DeviceKey.Name, DeviceKey.GetValue("FriendlyName"));
                                break;
                            }
                        }
                    }
                }
            }
            dataGridView1.Sort(dataGridView1.Columns[1], System.ComponentModel.ListSortDirection.Ascending);
        }
        private void dataGridView1_Leave(object sender, EventArgs e)
        {
            DataGridViewRow dataGridViewRow = (DataGridViewRow)sender;

        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell cell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
            string value = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            string key = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            Registry.SetValue(key, "FriendlyName", value);
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileDialog fileDialog = new SaveFileDialog();
            fileDialog.Filter = "CSV File|*.csv";
            fileDialog.Title = "Save a CSV File of Serial Key FriendlyName";
            fileDialog.ShowDialog();
            if (fileDialog.FileName != "")
            {
                System.IO.StreamWriter streamWriter = new System.IO.StreamWriter(fileDialog.FileName);
                try
                {
                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        streamWriter.WriteLine(dataGridView1.Rows[i].Cells[0].Value.ToString() + "," + dataGridView1.Rows[i].Cells[1].Value.ToString());
                    }
                    streamWriter.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "CSV File|*.csv";
            fileDialog.Title = "Load a CSV File of Serial Key FriendlyName";
            fileDialog.ShowDialog();
            if (fileDialog.FileName != "")
            {
                System.IO.StreamReader streamReader = new System.IO.StreamReader(fileDialog.FileName);
                try
                {
                    string[] lines = System.IO.File.ReadAllLines(fileDialog.FileName);
                    foreach (string line in lines)
                    {
                        string[] values = line.Split(',');
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (row.Cells[0].Value != null)
                            {
                                string key = row.Cells[0].Value.ToString();
                                if (key == values[0])
                                {
                                    row.Cells[1].Value = values[1];
                                    Registry.SetValue(key, "FriendlyName", values[1]);
                                    break;
                                }
                            }
                        }
                    }
                    streamReader.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
