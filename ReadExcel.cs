using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using IronXL;
using System.Diagnostics;
using System.Threading;

namespace ExcelReading
{
    public partial class ReadExcel : Form
    {
        public ReadExcel()
        {
            InitializeComponent();
            readExcel();
        }
        private void readExcel()
        {
            string filename = @"D:\Learning\WindowsExcelReading\ExcelReading\TestData.xlsx";
            DataTable dt = ReadExcelData(filename);
            for(int i = 0; i < dt.Rows.Count; i++)
            {
                // open in Internet Explorer
                Process.Start(@"C:\Program Files (x86)\Internet Explorer\iexplore.exe",
                  dt.Rows[i][0].ToString());
                Console.WriteLine(dt.Rows[i][0].ToString());
                /*Thread.Sleep(100);
                driver.Manage().Window.Maximize();
                Thread.Sleep(100);
                driver.Navigate().GoToUrl("https://automationlive.quality.pmapconnect.com/");
                //driver.Url = "https://automationlive.quality.pmapconnect.com/";
                Thread.Sleep(500);
                driver.FindElement(By.XPath("//span[text()='Launch product']")).Click();
                Thread.Sleep(5000);
                driver.FindElement(By.Id("txtUserName")).SendKeys("sateesh");
                driver.FindElement(By.Id("txtPassword")).SendKeys("sateesh");
                driver.FindElement(By.XPath("//input[@id='btnLogin']")).Click();
                Thread.Sleep(100);
                driver.Close();
                driver.Quit();
                */
            }

        }
        private DataTable ReadExcelData(string fileName)
        {
            WorkBook workbook = WorkBook.Load(fileName);
            WorkSheet sheet = workbook.DefaultWorkSheet;
            return sheet.ToDataTable(true);
        }
    }
}
