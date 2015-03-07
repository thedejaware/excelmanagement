using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace ExcelProcess
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnResults_Click(object sender, EventArgs e)
        {
            BindExcelValues();
        }

        private void RunMacro(object oApp, object[] oRunArgs)
        {
            oApp.GetType().InvokeMember("Run",
                System.Reflection.BindingFlags.Default |
                System.Reflection.BindingFlags.InvokeMethod,
                null, oApp, oRunArgs);
        }
        private void BindExcelValues()
        {
            try
            {
                string executableLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string xlsmLocation = Path.Combine(executableLocation, "MacroExcel.xlsm");

                //Run Macro
                object oMissing = System.Reflection.Missing.Value;
                Excel.ApplicationClass oExcel = new Excel.ApplicationClass();
                Excel.Workbooks oBooks = oExcel.Workbooks;
                Excel._Workbook oBook = null;
                //oExcel.Visible = true;
                oBook = oBooks.Open(xlsmLocation, oMissing, oMissing,
                    oMissing, oMissing, oMissing, oMissing, oMissing, oMissing,
                    oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

                // Run the macros.

                RunMacro(oExcel, new Object[] { "MacroExcel.xlsm!CustomExcelMacro" });
                // Quit Excel and clean up.               
                string minuteSecond = DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();
                string newXlsmLocation = Path.Combine(executableLocation, "MacroExcel" + minuteSecond + ".xlsm");
                oBook.SaveAs(new Object[] { newXlsmLocation });
                oBook.Close(false, oMissing, oMissing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook);
                oBook = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks);
                oBooks = null;
                oExcel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel);
                oExcel = null;

                string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + newXlsmLocation + "';Extended Properties=Excel 12.0;";
                OleDbConnection con = new OleDbConnection(connectionString);
                OleDbCommand com = new OleDbCommand("SELECT * from [Sheet1$]", con);
                OleDbDataAdapter adp = new OleDbDataAdapter(com);
                DataTable dt = new DataTable();
                adp.Fill(dt);
                grdExcel.AutoGenerateColumns = true;
                grdExcel.DataSource = dt;
            }
            catch (Exception exc)
            {

                MessageBox.Show(exc.InnerException.Message);
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
