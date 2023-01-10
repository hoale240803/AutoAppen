using AutoAppenWinform.Enum;
using AutoAppenWinform.Resources;
using AutoAppenWinform.Services.Interfaces;
using AutoAppenWinform.Utils;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace AutoAppenWinform
{
    public partial class AutoAppen : Form
    {
        private readonly IGmailService _gmailService;

        private const string clientID = "423901540646-ca1dl6hp64dd229t9uplnl1n1km4niv8.apps.googleusercontent.com";

        private const string clientSecret = "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";

        public AutoAppen(IGmailService gmailService)
        {
            InitializeComponent();
            _gmailService = gmailService;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _gmailService.SayHello();
        }

        private void loginEmailBtn_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("login email clicked");

            //_gmailService.GetUnreadMessages();

            _gmailService.LoginGmailAccount();
        }

        private void stopBtn_Click(object sender, EventArgs e)
        {
            MessageBox.Show("stop clicked");

            OpenBrowser();
        }

        private async void startBtn_Click(object sender, EventArgs e)
        {
            try
            {
                // TODO: 1. Validate at least one DataGrid selected
                if (dataGrid.SelectedRows.Count == 0 || dataGrid.SelectedRows.Count > 1)
                {
                    MessageBox.Show("Please select one Account");
                    return;
                }
                else if (AnyStartingAccount())
                {
                    MessageBox.Show(ErrorMessages.StartingAccount);
                    return;
                }

                // TODO: 2.1 Set state for Selected account

                SetStatusAccount();

                // TODO: 2.2 Open browser and login

               // OpenBrowser();

                string access_token = await _gmailService.GetAccessToken();
                _gmailService.GetGmailVerificationCode(access_token);

                // TODO: 3. Fake ip

                // TODO: 4. Register Appen

                // TODO: 5. Get verification code
                //

                // TODO: 6. Bypass Mobile Scan
            }
            catch (Exception)
            {
                throw;
            }
        }

        private bool AnyStartingAccount()
        {
            var rowFirstColumn = dataGrid.Rows;

            for (int i = 0; i < rowFirstColumn.Count; i++)
            {
                if (rowFirstColumn[i].Cells[0].Value == AccountStatusEnum.Starting.ToString())
                    return true;
            }

            return false;
        }

        private void SetStatusAccount()
        {
            // TODO: 1. Set Status for Selected Account

            var selectedIndex = dataGrid.SelectedRows[0].Index;

            dataGrid.Rows[selectedIndex].Cells[0].Value = AccountStatusEnum.Starting.ToString();
            var dataGridViewCellStyle = new DataGridViewCellStyle();
            dataGridViewCellStyle.BackColor = Color.GreenYellow;
            dataGridViewCellStyle.ForeColor = Color.Black;
            dataGridViewCellStyle.SelectionBackColor = Color.GreenYellow;
            dataGridViewCellStyle.SelectionForeColor = Color.Black;
            dataGrid.Rows[selectedIndex].Cells[0].Style = dataGridViewCellStyle;

            //MessageBox.Show($"Selected row is {selectedIndex}");

            // TODO: 2. Set Progress
        }

        private void OpenBrowser()
        {
            WebDriver driver = new ChromeDriver();
            try
            {
                // TODO: 1. Get authorization request
                var (authorizationReq, state, code_vefifier, code_challenge, redirectURI, http) = GetAuthorizationRequest();

                // TODO: 2. Navigation to authorizationReq;

                driver.Navigate().GoToUrl("https://mail.google.com/mail/");
                Thread.Sleep(2000);



                // TODO: Test screenshot 


                _gmailService.TakeScreenShot(driver);

                var emailTxtboxXPath = @"//*[@id=""identifierId""]";
                IWebElement emailTxtbox = driver.FindElement(By.XPath(emailTxtboxXPath));
                if (emailTxtbox != null)
                {
                    // Identify Google Email
                    var emailParam = "trungleo08241999@gmail.com";
                    var pass = "HoaLGoogle99123!@#";
                    emailTxtbox.SendKeys(emailParam);

                    // TODO: 4. Identify Continue Button
                    var continueBtnXPath = @"//*[@id=""identifierNext""]/div/button";
                    IWebElement continueBtn = driver.FindElement(By.XPath(continueBtnXPath));
                    continueBtn.Click();
                    Thread.Sleep(2000);

                    // TODO: 5. Fill password
                    var passTxtXPath = @"//*[@id=""password""]/div[1]/div/div[1]/input";
                    IWebElement passTxt = driver.FindElement(By.Name("Passwd"));
                    passTxt.SendKeys(pass);

                    // TODO: 6. Identify Continue Button
                    var nextBtnXPath = @"//*[@id=""passwordNext""]/div/button";
                    IWebElement nextBtn = driver.FindElement(By.XPath(nextBtnXPath));
                    nextBtn.Click();

                    Thread.Sleep(2000);
                }

                // TODO: 3. Get Access Token
                GetAccessToken(http, state, code_vefifier, redirectURI);



            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally { driver.Quit(); }
        }

        private (string authorReq, string state, string code_vefifier, string code_challenge, string redirectURI, HttpListener http) GetAuthorizationRequest()
        {
            // Generates state and PKCE values.
            // More details https://developers.google.com/identity/protocols/oauth2/native-app
            string state = GmailUtils.randomDataBase64url(32);
            string code_verifier = GmailUtils.randomDataBase64url(32);
            string code_challenge = GmailUtils.base64urlencodeNoPadding(GmailUtils.sha256(code_verifier));
            const string code_challenge_method = "S256";

            // Creates a redirect URI using an available port on the loopback address.
            string redirectURI = string.Format("http://{0}:{1}/", IPAddress.Loopback, GmailUtils.GetRandomUnusedPort());
            _gmailService.Output("redirect URI: " + redirectURI);

            // Creates an HttpListener to listen for requests on that redirect URI.
            var http = new HttpListener();
            http.Prefixes.Add(redirectURI);
            _gmailService.Output("Listening..");
            http.Start();

            // Creates the OAuth 2.0 authorization request.
            var authorizationRequest = _gmailService.CreateGmailOAth2AuthorizationRequest(redirectURI, state, code_challenge, code_challenge_method);

            return (authorizationRequest, state, code_verifier, code_challenge, redirectURI, http);
        }

        private string GetAccessToken(HttpListener http, string state, string code_verifier, string redirectURI)
        {
            var context = http.GetContextAsync().Result;

            // Brings this app back to the foreground.
            //this.Activate();

            // Sends an HTTP response to the browser.
            SendHTTPResponseToBrowser(http, context);

            // Checks for errors.
            if (context.Request.QueryString.Get("error") != null)
            {
                var error = string.Format("OAuth authorization error: {0}.", context.Request.QueryString.Get("error"));
                Output(error);
                return error;
            }
            if (context.Request.QueryString.Get("code") == null
                || context.Request.QueryString.Get("state") == null)
            {
                var error = "Malformed authorization response. " + context.Request.QueryString;
                Output(error);
                return error;
            }

            // extracts the code
            var code = context.Request.QueryString.Get("code");
            var incoming_state = context.Request.QueryString.Get("state");

            // Compares the receieved state to the expected value, to ensure that
            // this app made the request which resulted in authorization.
            if (incoming_state != state)
            {
                var error = string.Format("Received request with invalid state ({0})", incoming_state);
                Output(error);
                return error;
            }
            Output("Authorization code: " + code);

            // Starts the code exchange at the Token Endpoint.
            var accessToken = performCodeExchange(code, code_verifier, redirectURI).Result;

            return accessToken;
        }

        private async Task<string> performCodeExchange(string code, string code_verifier, string redirectURI)
        {
            Output("Exchanging code for tokens...");

            // builds the  request
            string tokenRequestURI = "https://www.googleapis.com/oauth2/v4/token";
            string tokenRequestBody = string.Format("code={0}&redirect_uri={1}&client_id={2}&code_verifier={3}&client_secret={4}&scope=&grant_type=authorization_code",
                code,
                Uri.EscapeDataString(redirectURI),
            clientID,
                code_verifier,
                clientSecret
                );

            // sends the request
            HttpWebRequest tokenRequest = (HttpWebRequest)WebRequest.Create(tokenRequestURI);
            tokenRequest.Method = "POST";
            tokenRequest.ContentType = "application/x-www-form-urlencoded";
            tokenRequest.Accept = "Accept=text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
            byte[] _byteVersion = Encoding.ASCII.GetBytes(tokenRequestBody);
            tokenRequest.ContentLength = _byteVersion.Length;
            Stream stream = tokenRequest.GetRequestStream();
            await stream.WriteAsync(_byteVersion, 0, _byteVersion.Length);
            stream.Close();

            try
            {
                // gets the response
                WebResponse tokenResponse = await tokenRequest.GetResponseAsync();
                using (StreamReader reader = new StreamReader(tokenResponse.GetResponseStream()))
                {
                    // reads response body
                    string responseText = await reader.ReadToEndAsync();
                    Output(responseText);

                    // converts to dictionary
                    Dictionary<string, string> tokenEndpointDecoded = JsonConvert.DeserializeObject<Dictionary<string, string>>(responseText);

                    string access_token = tokenEndpointDecoded["access_token"];
                    //userinfoCall(access_token);

                    return access_token;
                }
            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.ProtocolError)
                {
                    var response = ex.Response as HttpWebResponse;
                    if (response != null)
                    {
                        Output("HTTP: " + response.StatusCode);
                        using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                        {
                            // reads response body
                            string responseText = await reader.ReadToEndAsync();
                            Output(responseText);
                            return responseText;
                        }
                    }
                }
                return ex.Message;
            }
        }

        private void SendHTTPResponseToBrowser(HttpListener http, HttpListenerContext context)
        {
            var response = context.Response;
            string responseString = string.Format("<html><head><meta http-equiv='refresh' content='10;url=https://google.com'></head><body>Please return to the app.</body></html>");
            var buffer = Encoding.UTF8.GetBytes(responseString);
            response.ContentLength64 = buffer.Length;
            var responseOutput = response.OutputStream;
            Task responseTask = responseOutput.WriteAsync(buffer, 0, buffer.Length).ContinueWith((task) =>
            {
                responseOutput.Close();
                http.Stop();
                Console.WriteLine("HTTP server stopped.");
            });
        }

        private void RegisterRaterAccount()
        {
            // TODO: 1.Gmail Authentication

            // TODO: 2.Register Append account

            // TODO: 3.Login

            // TODO: 4.Setup onboard account
        }

        private void SetupOnboardAccount()
        {
            // TODO: 1.Select account language

            // TODO: 2.Demographic Data Survey

            // TODO: 3.Setup Location

            // TODO: 4.Setup Education

            // TODO: 5.Setup Prior Experience

            // TODO: 6.Add Phone number

            // TODO: 7.Click Save and submit profile

            // TODO: 8.SelectSmartPhoneQuestionare

            // TODO: 9.Select Additional Profile information
        }

        private void ProvideAdditionalInfo()
        {
            // TODO: 1.Select four "yes"

            // TODO: 2.Your level of familiarity with computers and the internet?
        }

        private void SelectSmartphoneQuestionare()
        {
            // TODO: 1.Answer question "Do you own an internet enabled smartphone?"

            // TODO: 2.Scan QR Code

            // TODO: 3.Click continue
        }

        private void ScanQRCode()
        {
            // TODO: 1.Get link

            // TODO: 2.Scan QA code
        }

        private void ClickSaveAndSubmit()
        {
            // TODO: 1.Click Save and submit

            // TODO: 2.Click Continue click
        }

        private void AddPhoneNumber()
        {
            // TODO: 1.Fill PhoneNumber

            // TODO: 2.Click next:Preview
        }

        private void SetupPriorExperience()
        {
            // TODO: 1.Check experience box

            // TODO: 2.Click next:Phone Number
        }

        private void SetupEducation()
        {
            // TODO: 1.Select Highest of education

            // TODO: 2.Select Linguistics qualification

            // TODO: 3.Click Next: Prior Experience
        }

        private void SetupLocation()
        {
            // TODO: 1.Fill Street Address

            // TODO: 2.Fill City

            // TODO: 3.Fill State of Province

            // TODO: 4.Select Residence history

            // TODO: 5.Click next: Education
        }

        private void DemographicDataSurvey()
        {
            // TODO: 1.Select Gender
            // TODO: 1.Select Date of birth
            // TODO: 2.Select Ethnicity
            // TODO: 3.Select Complexion
            // TODO: 4.Select "Are you considered disabled?"

            //TODO:  5Click Submit
        }

        private void SelectAccountLanguage()
        {
            // TODO: 1.Choosing Primary Language

            // TODO: 2.Select Your Language Region

            // TODO: 3.Click Continue
        }

        private void LoginAppen()
        {
            // TODO: 1.Fill gmail

            // TODO: 2.Fill password

            // TODO: 3.Click Login btn
        }

        private void RegisterAppendAccount()
        {
            // TODO: 1.Fill Form

            // TODO: 2.Verify Email

            // TODO: 3.Enter verification code
        }

        private void VerifyEmail()
        {
            // TODO: 1.Go to mail box

            // TODO: 2.Select message sent by Append

            // TODO: 3.Get verification code
        }

        private void FillAppendAccountForm()
        {
            // TODO: 1.Fill email field

            // TODO: 2.Fill password

            // TODO: 3.Fill first name, last name

            // TODO: 4.Select country of residence

            // TODO: 5.Check over 18 ages

            // TODO: 6.Click Create Account btn
        }

        private void outputBox_TextChanged(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// Appends the given string to the on-screen log, and the debug console.
        /// </summary>
        /// <param name="output">string to be appended</param>
        public void Output(string output)
        {
            outputBox.Text = outputBox.Text + output + Environment.NewLine;
            Console.WriteLine(output);
        }

        private void openFileBtn_Click(object sender, EventArgs e)
        {
            try
            {
                profileExcelFile.ShowDialog();
                profileExcelFile.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
                //profileExcelFile.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

                string fileName = profileExcelFile.FileName;

                if (profileExcelFile.ShowDialog() == DialogResult.OK)
                {
                    fileName = profileExcelFile.FileName;

                    // Show path of path
                    excelFileTxt.Text = fileName;
                    LoadExcelFile(fileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw ex;
            }
        }

        private void LoadExcelFile(string fileName)
        {
            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(fileName);
            Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            dataGrid.ColumnCount = colCount + 1;
            dataGrid.RowCount = rowCount;

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    //write the value to the Grid
                    var cell = xlRange.Cells[i + 1, j];
                    var cellValue = xlRange.Cells[i + 1, j].Value2;
                    // xlRange.Cells[i + 1, j]  to skip Header
                    if (xlRange.Cells[i + 1, j] != null && xlRange.Cells[i + 1, j].Value2 != null)
                    {
                        dataGrid.Rows[i - 1].Cells[j].Value = xlRange.Cells[i + 1, j].Value2.ToString();
                    }
                    // Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");

                    //add useful things here!
                }
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }
}