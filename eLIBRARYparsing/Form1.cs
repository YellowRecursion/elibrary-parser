using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
// Библиотеки для работы с excel
using Excel = Microsoft.Office.Interop.Excel;
using Keys = OpenQA.Selenium.Keys;

namespace eLIBRARYparsing
{
    public partial class Form1 : Form
    {
        string fileDirectory;
        string resultDirectory;

        bool started;
        Thread t;

        int k = 0;

        // Создаём экземпляр нашего приложения
        Excel.Application excelApp = new Excel.Application();
        // Создаём экземпляр рабочий книги Excel
        Excel.Workbook workBook;
        // Создаём экземпляр листа Excel
        Excel.Worksheet workSheet;

        string name = "Коледин Сергей Николаевич";

        //StreamWriter sw;


        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            resultDirectory = folderBrowserDialog1.SelectedPath;
            textBox2.Text = folderBrowserDialog1.SelectedPath;
        }

        private void OpenFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            fileDirectory = openFileDialog1.FileName;
            textBox1.Text = openFileDialog1.FileName;
        }

        private void FolderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void Button3_Click(object sender, EventArgs e)
        {
            if (!started)
            {
                button1.Enabled = false;
                button2.Enabled = false;
                button3.Text = "Stop";
                t = new Thread(Main);
                t.SetApartmentState(ApartmentState.STA);
                t.IsBackground = true;
                t.Start();
                print("Starting...");
            }
            else
            {
                button1.Enabled = true;
                button2.Enabled = true;
                button3.Text = "Start";
                if (t.IsAlive)
                {
                    t.Abort();
                    t.Interrupt();
                }
                print("Stoping...");
                workBook.Close();
                //if (sw != null) sw.Close();
            }

            started = !started;
        }



        void Main()
        {
            //try
            {
                Random r = new Random();

                workBook = excelApp.Workbooks.Open(fileDirectory);
                workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);

                ChromeOptions chromeOptions = new ChromeOptions();
                //chromeOptions.AddArgument("headless");
                IWebDriver driver = new ChromeDriver(chromeOptions);

                //FirefoxOptions chromeOptions = new FirefoxOptions();
                //IWebDriver driver = new FirefoxDriver(chromeOptions);

                Thread.Sleep(1000);
                driver.Url = @"https://www.elibrary.ru/";
                IWebElement element;

                //element = driver.FindElement(By.XPath(@".//input[@id='login']"));
                //element.SendKeys(textBox3.Text);

                //element = driver.FindElement(By.XPath(@".//input[@id='password']"));
                //element.SendKeys(textBox4.Text);

                element = driver.FindElement(By.XPath(@".//input[@id='login']"));
                //element.SendKeys("vadik00056");
                element.SendKeys(textBox3.Text);

                element = driver.FindElement(By.XPath(@".//input[@id='password']"));
                //element.SendKeys("vadim1232684");
                element.SendKeys(textBox4.Text);

                element = driver.FindElement(By.XPath(@".//div[@onclick='check_all()']"));
                element.Click();

                Thread.Sleep(3000);

                while (workSheet.Cells[k + 1, 1].Text.ToString() != "")
                {
                    //try
                    //{
                    name = workSheet.Cells[k + 1, 1].Text.ToString();
                    print((k + 1) + ". " + name);

                    if (!Directory.Exists(resultDirectory + @"\Text")) Directory.CreateDirectory(resultDirectory + @"\Text");
                    if (File.Exists(resultDirectory + @"\Text\" + name + ".txt"))
                    {
                        if (File.ReadAllText(resultDirectory + @"\Text\" + name + ".txt").Length > 10)
                        {
                            print("Информация уже имеется");
                            k++;
                            continue;
                        }
                    }

                    driver.Url = @"https://www.elibrary.ru/authors.asp";
                    Thread.Sleep(r.Next(2000, 7000));

                    element = driver.FindElement(By.XPath(@".//input[@id='surname']"));
                    element.Clear();
                    element.SendKeys(name);

                    element = driver.FindElement(By.XPath(@".//select[@name='town']"));
                    element.SendKeys("Уфа");

                    IJavaScriptExecutor js = driver as IJavaScriptExecutor;
                    js.ExecuteScript("author_search()");
                    //element = driver.FindElement(By.XPath(@".//div[@onclick='author_search()']"));
                    //Thread.Sleep(500);
                    //element.Click();
                    //Thread.Sleep(500);
                    //element.Click();

                    Thread.Sleep(500);

                    IWebElement man;
                    try
                    {
                        man = driver.FindElement(By.XPath(@".//tr[@bgcolor='#f5f5f5']"));
                    }
                    catch (Exception)
                    {
                        k++;
                        print("Не найдено авторов, удовлетворяющих условиям поиска");
                        continue;
                    }


                    Thread.Sleep(r.Next(100, 300));

                    int countOfArticles;
                    try
                    {
                        countOfArticles = Convert.ToInt32(man.FindElement(By.XPath(@".//a[@title='Список публикаций данного автора в РИНЦ']")).Text);
                        var elem = man.FindElement(By.XPath(@".//a[@title='Список публикаций данного автора в РИНЦ']"));
                        IJavaScriptExecutor executor = (IJavaScriptExecutor)driver;
                        executor.ExecuteScript("arguments[0].click();", elem);
                    }
                    catch (Exception)
                    {
                        k++;
                        print("Не найдена кнопка 'Список публикаций данного автора в РИНЦ'");
                        continue;
                    }

                    Thread.Sleep(1000);
                    Thread.Sleep(r.Next(100, 300));


                    IWebElement[] pages = driver.FindElements(By.XPath(@".//tr[@class='menurb']/td[@class='mouse-hovergr']")).ToArray();


                    Thread.Sleep(r.Next(100, 300));

                    for (int j = (countOfArticles < 21) ? -3 : 0; j < pages.Length - 2; j++) // ИСПРАВИТЬ БАГ ЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁЁ
                    {
                        print("PAGE #" + (j + 1) + "/" + (pages.Length - 2));
                        Thread.Sleep(r.Next(3000, 6000));
                        IWebElement[] articles = driver.FindElement(By.XPath(@".//table[@id='restab']")).FindElements(By.TagName("tr")).ToArray();
                        if (articles.Length > 0)
                        {
                            for (int i = 3; i < articles.Length; i++)
                            {
                                print("Article: " + (i - 2));
                                Thread.Sleep(r.Next(1000, 3000));

                                articles[i] = articles[i].FindElements(By.TagName("td"))[1].FindElement(By.TagName("a"));

                                // устанавливаем таймаут ожидания загрузки
                                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));

                                articles[i].SendKeys(Keys.Control + Keys.Return);
                                Thread.Sleep(1000);
                                try
                                {
                                    driver.SwitchTo().Window(driver.WindowHandles[1]);
                                    Thread.Sleep(r.Next(100, 300));
                                }
                                catch (TimeoutException)
                                {
                                    Thread.Sleep(2000);
                                }


                                Thread.Sleep(200);
                                Thread.Sleep(r.Next(100, 300));


                                // Copy info

                                IWebElement content;
                                IWebElement[] contents;

                                // Create directory, open files
                                if (!Directory.Exists(resultDirectory + @"\Text")) Directory.CreateDirectory(resultDirectory + @"\Text");

                                FileStream aFile = new FileStream(resultDirectory + @"\Text\" + name + ".txt", FileMode.OpenOrCreate);
                                StreamWriter sw = new StreamWriter(aFile);
                                aFile.Seek(0, SeekOrigin.End);

                                if (!Directory.Exists(resultDirectory + @"\UDC")) Directory.CreateDirectory(resultDirectory + @"\UDC");

                                FileStream bFile = new FileStream(resultDirectory + @"\UDC\" + name + ".txt", FileMode.OpenOrCreate);
                                StreamWriter sw2 = new StreamWriter(bFile);
                                bFile.Seek(0, SeekOrigin.End);

                                int c = 0;
                                while (!HasElement(driver, By.ClassName("bigtext")))
                                {
                                    c++;
                                    if (c > 8) break;
                                    Thread.Sleep(1000);
                                }

                                if (!HasElement(driver, By.ClassName("bigtext")))
                                {
                                    sw.Close();
                                    sw2.Close();
                                    print("bigtext not found");

                                    int h = 0;
                                    while (h < 10)
                                    {
                                        try
                                        {
                                            driver.Close();
                                            break;
                                        }
                                        catch (WebDriverException)
                                        {
                                            h++;
                                            Thread.Sleep(500);
                                            continue;
                                        }
                                    }


                                    Thread.Sleep(100);
                                    driver.SwitchTo().Window(driver.WindowHandles[0]);
                                    Thread.Sleep(200);
                                    continue;
                                }

                                content = driver.FindElement(By.ClassName("bigtext"));

                                // Text info copy
                                sw.WriteLine(content.Text);
                                contents = driver.FindElements(By.XPath(@".//table[@width='550']")).ToArray();
                                for (int o = 0; o < contents.Length; o++)
                                {
                                    if (contents[o].FindElement(By.TagName("tbody")).FindElements(By.TagName("tr"))[0].Text.Contains("КЛЮЧЕВЫЕ СЛОВА:"))
                                    {
                                        sw.WriteLine(contents[o].FindElement(By.TagName("tbody")).FindElements(By.TagName("tr"))[1].Text);
                                    }
                                    else
                                    if (contents[o].FindElement(By.TagName("tbody")).FindElements(By.TagName("tr"))[0].Text.Contains("АННОТАЦИЯ:"))
                                    {
                                        sw.WriteLine(contents[o].FindElement(By.TagName("tbody")).FindElements(By.TagName("tr"))[1].Text);
                                    }
                                    else
                                    if (contents[o].FindElement(By.TagName("tbody")).FindElements(By.TagName("tr"))[0].Text.Contains("КЛЮЧЕВЫЕ СЛОВА:"))
                                    {
                                        sw.WriteLine(contents[o].FindElement(By.TagName("tbody")).FindElements(By.TagName("tr"))[1].Text);
                                    }
                                }
                                // UDC copy
                                contents = driver.FindElements(By.XPath(@".//td[@width='574']")).ToArray();
                                for (int o = 0; o < contents.Length; o++)
                                {
                                    if (contents[o].Text.Contains("УДК:"))
                                    {
                                        sw2.WriteLine(contents[o].FindElement(By.TagName("font")).Text);
                                    }
                                }

                                sw.Close();
                                sw2.Close();

                                print("OK");


                                driver.Close();
                                Thread.Sleep(100);
                                driver.SwitchTo().Window(driver.WindowHandles[0]);
                                Thread.Sleep(200);
                            }
                        }
                        else
                        {
                            print("У автора нет статей");
                            Thread.Sleep(2000);
                        }

                        if (j < pages.Length - 1 - 2)
                        {
                            try
                            {
                                js.ExecuteScript("goto_page(" + (j + 2) + ")");
                            }
                            catch (Exception)
                            {
                                k++;
                                continue;
                            }

                        }
                        Thread.Sleep(1000);
                        Thread.Sleep(r.Next(100, 300));
                    }

                    k++;
                    //}
                    //catch
                    //{
                    //    k++;
                    //    continue;
                    //}
                    //sw.Close();
                }




                print("Completed!");

                workBook.Close();
                //button1.Enabled = true;
                //button2.Enabled = true;
                //button3.Text = "Запустить";
            }
            //catch
            {
                /*MessageBox.Show("Парсер экстренно завершил работу. Возможно сайт заблокировал доступ.");
                print("=COMPLETED=ERROR=");
                button1.Enabled = true; // ИСПРАВИТЬТ
                button2.Enabled = true;
                button3.Text = "Start";
                workBook.Close();*/
            }
        }

        public void print(object obj)
        {
            richTextBox1.Invoke(new Action(() => richTextBox1.Text = richTextBox1.Text + obj.ToString() + Environment.NewLine));
            richTextBox1.Invoke(new Action(() => richTextBox1.SelectionStart = richTextBox1.Text.Length));
            richTextBox1.Invoke(new Action(() => richTextBox1.ScrollToCaret()));
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (workBook != null)
                workBook.Close();
        }

        public bool HasElement(IWebDriver drv, By selector)
        {
            try
            {
                return drv.FindElements(selector).Count > 0;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool ShowWindow(IntPtr hWnd, int showWindowCommand);
    }
}
