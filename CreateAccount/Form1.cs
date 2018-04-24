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
    public class Info
    {
        public string firstName;
        public string lastName;
        public string email;
        public string emailPassword;
        public string applePassword;
        public Info(string _fName,string _lName, string _e,string _ePw, string _aPw)
        {
            firstName = _fName;
            lastName = _lName;
            email = _e;
            emailPassword = _ePw;
            applePassword = _aPw;
        }
    }

    public partial class Form1 : Form
    {
        public IWebDriver emailDriver;
        public IWebDriver appleDrive;
        public int timeout;//5s
        IList<Info> infos;
        int current;

        public Form1()
        {
            InitializeComponent();
            timeout = Convert.ToInt16(ConfigurationManager.AppSettings["waitTime"].ToString());
            current = 0;
            infos = new List<Info>();
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
            DataTable data = ParseExcelFile(fileName);
            data.Rows[0].Delete();
            data.AcceptChanges();
            foreach(DataRow row in data.Rows)
            {
                Info info = new Info(row[1].ToString(), row[0].ToString(), row[2].ToString(),
                    row[3].ToString(), row[4].ToString());
                infos.Add(info);
            }
            button1.Enabled = false;
        }

        private string OpenEmail(string email, string password)
        {
            emailDriver = new ChromeDriver();
            emailDriver.Url = "https://mail.google.com";
            Thread.Sleep(timeout);

            emailDriver.FindElement(By.Name("identifier")).SendKeys(email);
            emailDriver.FindElement(By.XPath("//*[@id='identifierId']")).SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(timeout);
            emailDriver.FindElement(By.Name("password")).SendKeys(password);
            emailDriver.FindElement(By.Name("password")).SendKeys(OpenQA.Selenium.Keys.Enter);

            Thread.Sleep(2*timeout);
            try
            {
                emailDriver.FindElement(By.XPath("//span[contains(text(),'Verify your Apple ID email address')]")).Click();
            }
            catch { }            
            Thread.Sleep(timeout);
            return emailDriver.FindElement(By.XPath("//td[contains(@class,'verification-code')]")).Text;
        }

        private void Close()
        {
            //Signout

            //Close
            emailDriver.Dispose();
            emailDriver.Close();
            appleDrive.Dispose();
            appleDrive.Close();
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
                appleDrive.FindElement(By.XPath("//*[@id='confirm-password-input']")).SendKeys(password);
                
                IList<IWebElement> selects = appleDrive.FindElements(By.TagName("select"));
                for(int i=1;i<selects.Count;i++)
                {
                    IList<IWebElement> options = selects[i].FindElements(By.TagName("option"));
                    options[i].Click();
                    appleDrive.FindElements(By.XPath("//input[@placeholder='answer']"))[i-1].
                        SendKeys(firstName+DateTime.Now.ToString("o"));
                }                
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

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                CreateAccountApple(infos[current].email, infos[current].firstName,
                    infos[current].lastName, infos[current].applePassword);
                Thread.Sleep(5*timeout);
                string result = OpenEmail(infos[current].email, infos[current].emailPassword);

                for(int i =0;i<6;i++)
                {
                    //*[@id="char0"]
                    string id = "char" + i.ToString();
                    appleDrive.FindElement(By.XPath("//input[@id='"+id.ToString() + "']"))
                        .SendKeys(result[i].ToString());                    
                }
                appleDrive.FindElement(By.XPath("//div[contains(text(),'Continue')]")).Click();
                Thread.Sleep(timeout);
                Close();
                current++;
            }
            catch { }
        }
    }
}
