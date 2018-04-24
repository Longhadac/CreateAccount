using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Configuration;

namespace CreateAccount
{
    public partial class Form1 : Form
    {
        public IWebDriver emailDriver;
        public IWebDriver appleDrive;
        public int timeout;//5s
        public Form1()
        {
            InitializeComponent();
            timeout = Convert.ToInt16(ConfigurationManager.AppSettings["waitTime"].ToString());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fileName = @"C:\Users\Longhd\Desktop\Demo1.xlsx";
            //OpenFileDialog openFile = new OpenFileDialog();
            //DialogResult file = openFile.ShowDialog();
            //if (file == DialogResult.OK)
            //{
            //    fileName = openFile.FileName;
            //}

            try
            {
                DataTable data = ParseExcelFile(fileName);
                data.Rows[0].Delete();
                data.AcceptChanges();

                if (data.Rows.Count > 0)
                    foreach (DataRow row in data.Rows)
                    {
                        if (string.IsNullOrEmpty(row[0].ToString()))
                            continue;
                        CreateAccountApple(row[2].ToString(), row[1].ToString(), row[0].ToString(), row[4].ToString());
                    }
            }
            catch { }
        }

        private void OpenEmail(string email, string password)
        {
            emailDriver = new ChromeDriver();
            emailDriver.Url = "https://mail.google.com";
            Thread.Sleep(timeout);

            emailDriver.FindElement(By.Name("identifier")).SendKeys(email);
            emailDriver.FindElement(By.XPath("//*[@id='identifierId']")).SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(timeout);
            emailDriver.FindElement(By.Name("password")).SendKeys(password);
            emailDriver.FindElement(By.Name("password")).SendKeys(OpenQA.Selenium.Keys.Enter);
        }

        private void CloseEmail()
        {
            //Signout

            //Close
            emailDriver.Dispose();
            emailDriver.Close();
        }

        private void CreateAccountApple(string email, string firstName, string lastName, string password)
        {
            try
            {


                appleDrive = new ChromeDriver();
                appleDrive.Url = "https://appleid.apple.com/account#!&page=create";
                Thread.Sleep(timeout);
                appleDrive.FindElement(By.XPath("//input[@placeholder='first name']")).SendKeys(firstName);
                appleDrive.FindElement(By.XPath("//input[@placeholder='last name']")).SendKeys(lastName);
                appleDrive.FindElement(By.XPath("//input[@placeholder='birthday']")).SendKeys("11/11/1990");
                appleDrive.FindElement(By.XPath("//input[@placeholder='name@example.com']")).SendKeys(email);
                //appleDrive.FindElement(By.XPath("//*[@id='name1524047451989 - 0']")).SendKeys(firstName);
                //appleDrive.FindElement(By.XPath("//*[@id='name1524047452002-0']")).SendKeys(lastName);
                //appleDrive.FindElement(By.Id("input - 1524046911076 - 1")).SendKeys("11/11/1990");                
                //appleDrive.FindElement(By.Id("input - 1524046911103 - 0")).SendKeys(email);
                //
                //appleDrive.FindElement(By.XPath("//*[@id='confirm - password - input']")).Click();
                appleDrive.FindElement(By.XPath("//*[@id='confirm-password-input']")).SendKeys(password);

                //var dropdown = new SelectElement(driver.findElement(By.id("designation")));
                //appleDrive.FindElement(By.PartialLinkText("security-questions-answers/div/div[1]")).Click();
                IList<IWebElement> selects = appleDrive.FindElements(By.TagName("select"));
                for(int i=1;i<selects.Count;i++)
                {
                    IList<IWebElement> options = selects[i].FindElements(By.TagName("option"));
                    options[i].Click();
                    appleDrive.FindElements(By.XPath("//input[@placeholder='answer']"))[i-1].SendKeys(firstName);
                    //IWebElement select = options[i].FindElement(By.TagName("option"));
                    //select.Click();
                }
                //IWebElement select = appleDrive.FindElement(By.TagName("select"));
                //IWebElement firstOption = select.FindElement(By.TagName("option"));

                //appleDrive.FindElement(By.XPath("//select[contains(@id, 'security-questions-answers/div/div[1]')]")).Click();
                appleDrive.FindElements(By.XPath("//input[@placeholder='answer']"))[0].SendKeys(firstName);
                //appleDrive.FindElement(By.XPath("//*[@id='idms-step-1524130815766-0']/div[2]/div/div/div[4]/div/div/div/security-questions-answers/div/div[1]/security-question/div/div[1]/select")).Click();
                //appleDrive.FindElement(By.XPath("//*[@id='idms-step-1524130815766-0']/div[2]/div/div/div[4]/div/div/div/security-questions-answers/div/div[2]/security-question/div/div[1]/select")).Click();
                //appleDrive.FindElement(By.XPath("//*[@id='idms-step-1524130815766-0']/div[2]/div/div/div[4]/div/div/div/security-questions-answers/div/div[3]/security-question/div/div[1]/select")).Click();
                
                appleDrive.FindElement(By.XPath("//*[@id='password']")).SendKeys(password + OpenQA.Selenium.Keys.Tab);
            }
            catch(Exception ex)
            { }
        }

        private DataTable ParseExcelFile(string fileName)
        {
            DataTable results = new DataTable();
            string sheetName = ConfigurationManager.AppSettings["SheetName"];

            try
            {
                string connString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=no'", fileName);
                string sql = "SELECT * FROM [" + sheetName.ToString() + "]";
                using (OleDbConnection conn = new OleDbConnection(connString))
                {
                    conn.Open();
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        using (OleDbDataReader rdr = cmd.ExecuteReader())
                        {
                            results.Load(rdr);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Couldnot parse file");
            }
            return results;
        }
    }
}
