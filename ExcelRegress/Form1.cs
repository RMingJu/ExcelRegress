using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelRegress
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExecuteMacro_Regression(@"D:\RR_M.xlsm");
            //ExecuteMacro(@"D:\ATPVBAEN.XLAM");
            //ExecuteMacro_Regression(@"D:\ATPVBAEN.XLAM");
            


        }


        public static void ExecuteMacro_Regression(String filePath)
        {
            //    using Excel = Microsoft.Office.Interop.Excel;
            //    using System;
            //    using System.IO;


            
            var ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow;
            var wb = ExcelApp.Workbooks.Open(filePath, ReadOnly: false);

            
            try
            {
                ExcelApp.DisplayAlerts = false; //關閉通知
                ExcelApp.Visible = true;




                //        Application.Run "ATPVBAEN.XLAM!Regress", ActiveSheet.Range("$C$1:$C$61"), _
                //ActiveSheet.Range("$D$1:$K$61"), False, True, 95, "Test", False, False, _
                //False, False, , False
                //ExcelApp.Run("Test1");


                //foreach (Microsoft.Office.Interop.Excel.AddIn addIn in ExcelApp.AddIns)
                //{
                //    try
                //    {
                //        string s = addIn.Name;
                //        string path_add = addIn.Path;
                //        addIn.Installed = true;

                //    }
                //    catch
                //    {

                //    }
                //}



                //var bln = false;
                //bln = ExcelApp.RegisterXLL(@"C:\Program Files (x86)\Microsoft Office\Office12\Library\Analysis\ANALYS32.XLL");
                //bln = ExcelApp.RegisterXLL(@"C:\Program Files (x86)\Microsoft Office\Office12\Library\Analysis\FUNCRES.XLAM");
                //bln = ExcelApp.RegisterXLL(@"C:\Program Files (x86)\Microsoft Office\Office12\Library\Analysis\ATPVBATC.XLAM");
                //bln = ExcelApp.RegisterXLL(@"C:\Program Files (x86)\Microsoft Office\Office12\Library\Analysis\PROCDB.XLAM");
                //bln = ExcelApp.RegisterXLL(@"C:\Program Files (x86)\Microsoft Office\Office12\Library\Analysis\ATPVBAEN.XLAM");


                //=========重啟addIn===========================//
                foreach (Microsoft.Office.Interop.Excel.AddIn addIn in ExcelApp.AddIns)
                {
                    try
                    {
                        addIn.Installed = false;
                        addIn.Installed = true;
                    }
                    catch
                    {

                    }
                }
                //=========重啟addIn===========================//




                ExcelApp.Run("RunR","61");


             


                string fileName = DateTime.Now.ToString("yyyyyMMddHHmmss");
                wb.SaveAs( String.Format(@"D:\{0}.xlsm", fileName));


            }
            catch (Exception ex)
            {
                var test = ex.ToString();
            }
            finally {
                wb.Close(true, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                ExcelApp.Application.Quit();
                ExcelApp.Quit();

                
                ReleaseComObject(wb);
                ReleaseComObject(ExcelApp);


            }
            
        }


        private static void ReleaseComObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
        }

    }
}
