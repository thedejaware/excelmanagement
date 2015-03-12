using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelProcess
{
    public partial class Form2 : Form
    {
        System.Threading.Thread tPopup;
        System.Threading.Thread tSubmit;
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        #region Methods
        private void GetPointsExcelData()
        {
            string executableLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string xlsmLocation = Path.Combine(executableLocation, "ZiyaretSaati_Master.xlsm");

            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + xlsmLocation + "';Extended Properties=Excel 12.0;";
            OleDbConnection con = new OleDbConnection(connectionString);
            OleDbCommand com = new OleDbCommand("SELECT * from [Noktalar$]", con);
            OleDbDataAdapter adp = new OleDbDataAdapter(com);
            DataTable dt = new DataTable();
            adp.Fill(dt);
            grdPoints.DataSource = dt;
        }

        private void SubmitExcelDataToAccess(string pointName, string pointType, int pointCapacity, int pointStartAmount, int pointVisitAmount, int recycle, int dependentId, string weekendPgm, DateTime visitingTime, int visitDay1, int visitDay2, int visitDay3, int visitDay4, int visitDay5, int visitDay6, int visitDay7)
        {
            string executableLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string xlsmLocation = Path.Combine(executableLocation, "db.mdb");
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + xlsmLocation + "';Persist Security Info=False;";
            OleDbConnection con = new OleDbConnection(connectionString);
            OleDbCommand com = new OleDbCommand("INSERT INTO Noktalar(pointName,pointType,pointCapacity,pointStartAmount,pointVisitAmount,Recycle,dependentId,weekendPgm,visitingTime,VisitDay1,VisitDay2,VisitDay3,VisitDay4,VisitDay5,VisitDay6,VisitDay7) VALUES (@pointName, @pointType, @pointCapacity, @pointStartAmount, @pointVisitAmount, @recycle, @dependentId, @weekendPgm, @visitingTime, @visitDay1,@visitDay2,@visitDay3,@visitDay4,@visitDay5,@visitDay6,@visitDay7)", con);
            com.Parameters.AddWithValue("@pointName", pointName);
            com.Parameters.AddWithValue("@pointType", pointType);
            com.Parameters.AddWithValue("@pointCapacity", pointCapacity);
            com.Parameters.AddWithValue("@pointStartAmount", pointStartAmount);
            com.Parameters.AddWithValue("@pointVisitAmount", pointVisitAmount);
            com.Parameters.AddWithValue("@recycle", recycle);
            com.Parameters.AddWithValue("@dependentId", dependentId);
            com.Parameters.AddWithValue("@weekendPgm", weekendPgm);
            com.Parameters.AddWithValue("@visitingTime", visitingTime);
            com.Parameters.AddWithValue("@visitDay1", visitDay1);
            com.Parameters.AddWithValue("@visitDay2", visitDay2);
            com.Parameters.AddWithValue("@visitDay3", visitDay3);
            com.Parameters.AddWithValue("@visitDay4", visitDay4);
            com.Parameters.AddWithValue("@visitDay5", visitDay5);
            com.Parameters.AddWithValue("@visitDay6", visitDay6);
            com.Parameters.AddWithValue("@visitDay7", visitDay7);
            if (con.State == ConnectionState.Closed)
                con.Open();
            com.ExecuteNonQuery();
            if (con.State == ConnectionState.Open)
                con.Close();

        }

        private void CreateComboboxColumn(string[] dataPropertyName)
        {
            for (int i = 0; i < dataPropertyName.Length; i++)
            {
                DataGridViewComboBoxColumn columnName = new DataGridViewComboBoxColumn();
                List<ComboDataCS> mList = new List<ComboDataCS>();
                mList.Add(new ComboDataCS("Evet", 1.0));
                mList.Add(new ComboDataCS("Hayır", 0.0));
                columnName.HeaderText = dataPropertyName[i];
                columnName.Name = dataPropertyName[i];
                columnName.Width = 120;
                columnName.DataPropertyName = dataPropertyName[i];
                columnName.DataSource = mList;
                columnName.ValueMember = "ComboValue";
                columnName.DisplayMember = "ComboText";
                columnName.ValueType = grdPoints.Columns[dataPropertyName[i]].ValueType;
                columnName.DisplayIndex = grdPoints.Columns[dataPropertyName[i]].Index;
                grdPoints.Columns.Add(columnName);
            }
            for (int j = 0; j < dataPropertyName.Length; j++)
            {
                grdPoints.Columns.RemoveAt(grdPoints.Columns[dataPropertyName[j]].Index);
            }
        }

        private void ShowPopup()
        {
            Form2.CheckForIllegalCrossThreadCalls = false;
            var popup = new FormPopup();
            popup.ShowDialog();
        }

        private void SubmitAccess()
        {
            foreach (DataGridViewRow dataRow in grdPoints.Rows)
            {
                if (dataRow.DataBoundItem != null)
                {
                    SubmitExcelDataToAccess(dataRow.Cells["İsmi"].Value.ToString(), dataRow.Cells["Tipi"].Value.ToString(), int.Parse(dataRow.Cells["Kapasitesi"].Value.ToString()), int.Parse(dataRow.Cells["Başlangıç Para Miktari"].Value.ToString()), int.Parse(dataRow.Cells["Ziyarette Birakilacak Talep (Gün)"].Value.ToString()), int.Parse(dataRow.Cells["Recycle?"].Value.ToString()), int.Parse(dataRow.Cells["Bağlı Nokta ID"].Value.ToString()), dataRow.Cells["HaftaSonu PGM ye bağlı?"].Value.ToString(), Convert.ToDateTime(dataRow.Cells["Yaklaşık Ziyaret Zamanı"].Value.ToString()), int.Parse(dataRow.Cells["Gün 1 Ziyaret Edilebilir?"].Value.ToString()), int.Parse(dataRow.Cells["Gün 2 Ziyaret Edilebilir?"].Value.ToString()), int.Parse(dataRow.Cells["Gün 3 Ziyaret Edilebilir?"].Value.ToString()), int.Parse(dataRow.Cells["Gün 4 Ziyaret Edilebilir?"].Value.ToString()), int.Parse(dataRow.Cells["Gün 5 Ziyaret Edilebilir?"].Value.ToString()), int.Parse(dataRow.Cells["Gün 6 Ziyaret Edilebilir?"].Value.ToString()), int.Parse(dataRow.Cells["Gün 7 Ziyaret Edilebilir?"].Value.ToString()));
                }
                else
                {
                    tPopup.Abort();                 
                    MessageBox.Show("Aktarma işlemi başarıyla gerçekleşmiştir.");
                    tSubmit.Abort();
                    break;
                }

            }
        }
        #endregion

        #region Events
        private void btnExcel_Click(object sender, EventArgs e)
        {
            try
            {
                GetPointsExcelData();
                CreateComboboxColumn(new string[] { "Gün 1 Ziyaret Edilebilir?", "Gün 2 Ziyaret Edilebilir?", "Gün 3 Ziyaret Edilebilir?", "Gün 4 Ziyaret Edilebilir?", "Gün 5 Ziyaret Edilebilir?", "Gün 6 Ziyaret Edilebilir?", "Gün 7 Ziyaret Edilebilir?" });
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void btnAccess_Click(object sender, EventArgs e)
        {

            try
            {
                tPopup = new System.Threading.Thread(new System.Threading.ThreadStart(ShowPopup));
                tPopup.Start();
                tSubmit = new System.Threading.Thread(new System.Threading.ThreadStart(SubmitAccess));
                tSubmit.Start();

            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }

        }

        #endregion
    }
}
