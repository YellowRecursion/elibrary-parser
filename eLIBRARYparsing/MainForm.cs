using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
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
    public enum WorkStatus
    {
        NotProcessed,
        OK,
        Interrupted,
        Error
    }

    public partial class MainForm : Form
    {
        private const string URL_ELIBRARY = "https://www.elibrary.ru/";
        private const string URL_ELIBRARY_AUTHORS = "https://www.elibrary.ru/authors.asp";

        private string _workFilePath; // Excel file
        private string _resultsDirectory; // Folder for results

        // Input excel file offsets:
        private Vector2Int _namesOffset = new Vector2Int(1, 5);
        private Vector2Int _cityOffset = new Vector2Int(2, 5);
        private Vector2Int _countryOffset = new Vector2Int(3, 5);
        private Vector2Int _organizationOffset = new Vector2Int(4, 5);
        private Vector2Int _statusOffset = new Vector2Int(6, 5);
        private Vector2Int _pageOffset = new Vector2Int(7, 5);
        private Vector2Int _articleNumberOffset = new Vector2Int(8, 5);

        private bool _isStarted;
        private Thread _workThread;
        private Random _random;

        private Excel.Application _excelApp = new Excel.Application();
        private Excel.Workbook _workBook;
        private Excel.Worksheet _workSheet;



        public string AuthorName => GetCellValue(_namesOffset).ToString();
        public string AuthorCity => GetCellValue(_cityOffset).ToString();
        public string AuthorCountry => GetCellValue(_countryOffset).ToString();
        public string AuthorOrganization => GetCellValue(_organizationOffset).ToString();
        public WorkStatus AuthorWorkStatus
        {
            get
            {
                if (IsCellEmpty(_statusOffset)) return WorkStatus.NotProcessed;
                return (WorkStatus)(int)GetCellValue(_statusOffset);
            }
            set
            {
                SetCellValue(_statusOffset, value);
            }
        }
        public int AuthorPage
        {
            get
            {
                if (IsCellEmpty(_pageOffset)) return 1;
                return (int)GetCellValue(_pageOffset);
            }
            set
            {
                SetCellValue(_pageOffset, value);
            }
        }
        public int AuthorArticleNumber
        {
            get
            {
                if (IsCellEmpty(_articleNumberOffset)) return 0;
                return (int)GetCellValue(_articleNumberOffset);
            }
            set
            {
                SetCellValue(_articleNumberOffset, value);
            }
        }



        // INITIALIZING
        public MainForm()
        {
            InitializeComponent();
            _random = new Random();
        }



        // INPUT AND OUTPUT PREPARATION
        private void OpenExcelFileDialog(object sender, EventArgs e)
        {
            openExcelFileDialog.ShowDialog();
        }
        private void ExcelFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            _workFilePath = openExcelFileDialog.FileName;
            inputPathTextBox.Text = openExcelFileDialog.FileName;
        }
        private void CreateAndSelectOutputDirectory()
        {
            _resultsDirectory = Path.Combine(Path.GetDirectoryName(_workFilePath), $"{Path.GetFileNameWithoutExtension(_workFilePath)} Results");

            if (!Directory.Exists(_resultsDirectory))
            {
                Directory.CreateDirectory(_resultsDirectory);
                Directory.CreateDirectory(Path.Combine(_resultsDirectory, "Text"));
                Directory.CreateDirectory(Path.Combine(_resultsDirectory, "UDC"));
            }
        }



        // START
        private void StartButton_Click(object sender, EventArgs e)
        {
            CreateAndSelectOutputDirectory();
            Start();
        }
        private void Start()
        {
            if (!_isStarted)
            {
                // START

                selectInputPathButton.Enabled = false;
                startButton.Text = "Stop";
                _workThread = new Thread(Main);
                _workThread.SetApartmentState(ApartmentState.STA);
                _workThread.IsBackground = true;
                _workThread.Start();
                Print("Starting...");
            }
            else
            {
                // STOP

                selectInputPathButton.Enabled = true;
                startButton.Text = "Start";
                if (_workThread.IsAlive)
                {
                    _workThread.Abort();
                    _workThread.Interrupt();
                }
                Print("Stoping...");
                _workBook.Close();
            }

            _isStarted = !_isStarted;
        }



        // WORK BLOCK
        private IWebDriver _driver;
        private IJavaScriptExecutor JavaScriptExecutor => (IJavaScriptExecutor)_driver;
        private int _currentDataIndex;
        private void CreateDriver()
        {
            if (_driver != null) _driver.Close();

            _driver = new ChromeDriver();
            _driver.Url = URL_ELIBRARY;

            WaitForPassRobotTest();
        }
        private void AuthrizeOnElibrary()
        {
            if (_driver.Url != URL_ELIBRARY) _driver.Url = URL_ELIBRARY;

            var element = _driver.FindElement(By.XPath(@".//input[@id='login']"));
            element.SendKeys(loginField.Text);
            element = _driver.FindElement(By.XPath(@".//input[@id='password']"));
            element.SendKeys(passwordField.Text);
            element = _driver.FindElement(By.XPath(@".//div[@onclick='check_all()']"));
            element.Click();

            WaitForPassRobotTest();
        }
        private void Main()
        {
            _workBook = _excelApp.Workbooks.Open(_workFilePath);
            _workSheet = (Excel.Worksheet)_workBook.Worksheets.get_Item(1);

            _currentDataIndex = 0;

            while (!IsCellEmpty(_namesOffset))
            {
                Print($"{_currentDataIndex + 1}) {AuthorName}");

                switch (AuthorWorkStatus)
                {
                    case WorkStatus.OK:
                        Print("OK");
                        break;
                    default:
                        Work();
                        break;
                }
            }

            Print("Сompleted!");
        }
        private void Work()
        {
            AuthorWorkStatus = WorkStatus.Interrupted;

            CreateDriver();
            // AuthrizeOnElibrary() <- не факт что вообще нужно

            SleepInaccurateTime(3000);

            // Поиск автора
            try
            {
                FindAuthor();
            }
            catch (Exception ex)
            {
                Print("Возника ошибка при поиске автора");
                Print(ex.Message);
                AuthorWorkStatus = WorkStatus.Error;
                return;
            }
            if (!WebElementExists(By.XPath(@".//tr[@bgcolor='#f5f5f5']")))
            {
                Print("Не найдено авторов, удовлетворяющих условиям поиска");
                AuthorWorkStatus = WorkStatus.Error;
                return;
            }


            // Выбор автора из списка и открытие страницы автора
            try
            {
                var element = _driver.FindElement(By.XPath(@".//tr[@bgcolor='#f5f5f5']//a[@title='Список публикаций данного автора в РИНЦ']"));
                JavaScriptExecutor.ExecuteScript("arguments[0].click();", element);
            }
            catch (Exception)
            {
                Print("Не удалось нажать на кнопку 'Список публикаций данного автора в РИНЦ'");
                AuthorWorkStatus = WorkStatus.Error;
                return;
            }

            WaitForPassRobotTest();

            SleepInaccurateTime(2000);

            while (true)
            {
                JavaScriptExecutor.ExecuteScript($"goto_page({AuthorPage})");

                WaitForPassRobotTest();

                SleepInaccurateTime(2000);

                var articles = _driver.FindElements(By.TagName(".//table[@id='restab']//tr[@valign='middle'][@id]")).ToArray();

                if (articles.Length == 0)
                {
                    Print("OK");
                    AuthorWorkStatus = WorkStatus.OK;
                    return;
                }

                Print($"Page: {AuthorPage}");

                for (int i = AuthorArticleNumber; i < articles.Length; i++)
                {
                    Print($"Article: {i} / {articles.Length}");

                    // Click 
                    articles[i].FindElement(By.XPath("//a[@href]")).Click();

                    WaitForPassRobotTest();

                    if (WaitForWebElementNotExists(By.ClassName("bigtext"), 10))
                    {
                        SleepInaccurateTime(3000);

                        ProcessArticle();

                        Print("OK");
                    }
                    else
                    {
                        Print("Error: article name is not found");
                    }

                    // Back to articles page
                    _driver.Navigate().Back();
                    SleepInaccurateTime(2000);
                    articles = _driver.FindElements(By.TagName(".//table[@id='restab']//tr[@valign='middle'][@id]")).ToArray();

                    AuthorArticleNumber = i;
                }

                AuthorPage++;
                AuthorArticleNumber = 0;
            }
        }
        private void FindAuthor()
        {
            // _driver.Url = URL_ELIBRARY_AUTHORS;
            _driver.FindElement(By.XPath(@".//a[@href='/authors.asp']")).Click();

            WaitForPassRobotTest();

            SleepInaccurateTime(4000);

            var element = _driver.FindElement(By.XPath(@".//input[@id='surname']"));
            element.Clear();
            element.SendKeys(AuthorName);

            SleepInaccurateTime(1500);

            element = _driver.FindElement(By.XPath(@".//select[@name='town']"));
            element.SendKeys(AuthorCity);

            SleepInaccurateTime(1500);

            IJavaScriptExecutor js = _driver as IJavaScriptExecutor;
            js.ExecuteScript("author_search()");

            WaitForPassRobotTest();

            SleepInaccurateTime(2000);
        }
        private void ProcessArticle()
        {
            StreamWriter textFile = new StreamWriter(Path.Combine(_resultsDirectory, "Text", $"{AuthorName}.txt"), true);

            // Article name
            if (WebElementExists(By.XPath(".//p[@class='bigtext']"), out var webElement))
            {
                textFile.WriteLine(webElement.Text);
                Print("- название: +");
            }
            else
            {
                Print("- название: -");
            }

            // Article keywords
            if (WebElementExists(By.XPath(".//table[@width='550'][@cellpadding='3'][contains(., 'КЛЮЧЕВЫЕ СЛОВА:')]//tr[2]"), out webElement))
            {
                textFile.WriteLine(webElement.Text);
                Print("- ключевые слова: +");
            }
            else
            {
                Print("- ключевые слова: -");
            }

            // Article annotation
            if (WebElementExists(By.XPath(".//table[@width='550'][@cellpadding='3'][contains(., 'АННОТАЦИЯ:')]//tr[2]"), out webElement))
            {
                textFile.WriteLine(webElement.Text);
                Print("- аннотация: +");
            }
            else
            {
                Print("- аннотация: -");
            }

            textFile.Close();

            // Article UDC
            if (WebElementExists(By.XPath(".//td[@width='574'][contains(., 'УДК:')]/font")))
            {
                StreamWriter udcFile = new StreamWriter(Path.Combine(_resultsDirectory, "UDC", $"{AuthorName}.txt"), true);
                udcFile.WriteLine(_driver.FindElement(By.XPath(".//td[@width='574'][contains(., 'УДК:')]/font")).Text);
                udcFile.Close();
                Print("- УДК: +");
            }
            else
            {
                Print("- УДК: -");
            }
        }
        private void WaitForPassRobotTest()
        {
            while (WebElementExists(By.XPath(".//div[@class='midtext'][contains(., 'С Вашего IP-адреса')]")))
            {
                Thread.Sleep(3000);
            }
        }



        // FINISH WORK
        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (_workBook != null)
                _workBook.Close();
        }



        // EXCEL FILE WORKING UTILITIES
        private bool IsCellEmpty(Vector2Int offset, int index)
        {
            return GetCellValue(offset, index).ToString() == string.Empty;
        }
        private bool IsCellEmpty(Vector2Int offset)
        {
            return GetCellValue(offset).ToString() == string.Empty;
        }
        private object GetCellValue(Vector2Int offset, int index)
        {
            return _workSheet.Cells[index + offset.Y, offset.X].Text;
        }
        private object GetCellValue(Vector2Int offset)
        {
            return _workSheet.Cells[_currentDataIndex + offset.Y, offset.X].Text;
        }
        private void SetCellValue(Vector2Int offset, int index, object value)
        {
            _workSheet.Cells[index + offset.Y, offset.X].Text = value;
        }
        private void SetCellValue(Vector2Int offset, object value)
        {
            _workSheet.Cells[_currentDataIndex + offset.Y, offset.X].Text = value;
        }



        // PARSING UTILITIES
        private bool WebElementExists(By selector)
        {
            return _driver.FindElements(selector).Count > 0;
        }
        private bool WebElementExists(By selector, out IWebElement webElement)
        {
            var elements = _driver.FindElements(selector);

            if (elements.Count > 0) webElement = elements[0];
            else webElement = null;

            return elements.Count > 0;
        }
        /// <param name="maxSeconds">Max time to wait</param>
        /// <returns>True if element is exists</returns>
        private bool WaitForWebElementNotExists(By selector, int maxSeconds)
        {
            int timer = 0;

            while (!WebElementExists(selector))
            {
                Thread.Sleep(1000);

                timer++;

                if (timer >= maxSeconds)
                {
                    return false;
                }
            }

            return true;
        }



        // TOOLS
        private void Print(object obj)
        {
            logs.Invoke(new Action(() =>
            {
                logs.Text = logs.Text + obj.ToString() + Environment.NewLine;
                logs.SelectionStart = logs.Text.Length;
                logs.ScrollToCaret();
            }));
        }
        /// <summary>
        /// Возвращает случайное число в диапазоне [n - 25%, n + 25%)
        /// </summary>
        private int GetInaccurateNumber(int n)
        {
            return _random.Next(n - (int)(n * 0.25f), n + (int)(n * 0.25f));
        }
        private void SleepInaccurateTime(int targetMillisecondsTimeout)
        {
            Thread.Sleep(GetInaccurateNumber(targetMillisecondsTimeout));
        }



        // DPI FIX
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool ShowWindow(IntPtr hWnd, int showWindowCommand);
    }
}
