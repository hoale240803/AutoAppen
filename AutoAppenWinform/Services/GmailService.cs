using AutoAppenWinform.Models.HideMyAcc;
using AutoAppenWinform.Utils;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.VisualBasic.ApplicationServices;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using ZXing;
using ZXing.Common;
using static AutoAppenWinform.Services.StartProfileOptions;
using GmailServices = Google.Apis.Gmail.v1.GmailService;
using IGmailService = AutoAppenWinform.Services.Interfaces.IGmailService;
using Message = Google.Apis.Gmail.v1.Data.Message;
using Process = System.Diagnostics.Process;

namespace AutoAppenWinform.Services
{
    public class GmailService : IGmailService
    {
        // https://laptrinhvb.net/bai-viet/chuyen-de-csharp/---Csharp----Huong-dan-doc-gmail-su-dung-Gmail-API-lap-trinh-Csharp/ee736484b32e9cff.html
        private readonly GmailServices _gmailServiceCon;

        //public GmailService(GmailServices gmailServiceCon)
        //{
        //    _gmailServiceCon = GetGmailServiceAsync().Result;
        //}

        //private const string clientID = "423901540646-8l4jp8f1vggqunai7dopgjp2gj77nkib.apps.googleusercontent.com";

        //private const string clientSecret = "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";

        private const string clientID = "423901540646-ca1dl6hp64dd229t9uplnl1n1km4niv8.apps.googleusercontent.com";

        private const string clientSecret = "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";

        private const string authorizationEndpoint = "https://accounts.google.com/o/oauth2/v2/auth";

        public async void LoginGmailAccount()
        {
            // Generates state and PKCE values.
            // More details https://developers.google.com/identity/protocols/oauth2/native-app
            string state = GmailUtils.randomDataBase64url(32);
            string code_verifier = GmailUtils.randomDataBase64url(32);
            string code_challenge = GmailUtils.base64urlencodeNoPadding(GmailUtils.sha256(code_verifier));
            const string code_challenge_method = "S256";

            // Creates a redirect URI using an available port on the loopback address.
            string redirectURI = string.Format("http://{0}:{1}/", IPAddress.Loopback, GmailUtils.GetRandomUnusedPort());
            Output("redirect URI: " + redirectURI);

            // Creates an HttpListener to listen for requests on that redirect URI.
            var http = new HttpListener();
            http.Prefixes.Add(redirectURI);
            Output("Listening..");
            http.Start();

            // Creates the OAuth 2.0 authorization request.
            var authorizationRequest = CreateOAth2AuthorizationRequest(redirectURI, state, code_challenge, code_challenge_method);

            // Opens request in the browser.

            // TODO: Debug in library in C#

            try
            {
                Process.Start("\"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe\"", authorizationRequest);
            }
            catch
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Process.Start(new ProcessStartInfo("cmd", $"/c start {authorizationRequest}") { CreateNoWindow = true });

                    //new ProcessStartInfo("msedge")
                    //{
                    //    UseShellExecute = true,
                    //    Arguments = authorizationRequest
                    //};
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    Process.Start("xdg-open", authorizationRequest);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    Process.Start("open", authorizationRequest);
                }
                else
                {
                    throw;
                }
            }

            // Waits for the OAuth authorization response.
            var context = await http.GetContextAsync();

            // Brings this app back to the foreground.
            //this.Activate();

            // Sends an HTTP response to the browser.
            SendHTTPResponseToBrowser(http, context);

            // Checks for errors.
            if (context.Request.QueryString.Get("error") != null)
            {
                Output(string.Format("OAuth authorization error: {0}.", context.Request.QueryString.Get("error")));
                return;
            }
            if (context.Request.QueryString.Get("code") == null
                || context.Request.QueryString.Get("state") == null)
            {
                Output("Malformed authorization response. " + context.Request.QueryString);
                return;
            }

            // extracts the code
            var code = context.Request.QueryString.Get("code");
            var incoming_state = context.Request.QueryString.Get("state");

            // Compares the receieved state to the expected value, to ensure that
            // this app made the request which resulted in authorization.
            if (incoming_state != state)
            {
                Output(string.Format("Received request with invalid state ({0})", incoming_state));
                return;
            }
            Output("Authorization code: " + code);

            // Starts the code exchange at the Token Endpoint.
            performCodeExchange(code, code_verifier, redirectURI);
        }

        private string CreateOAth2AuthorizationRequest(string redirectURI, string state, string code_challenge, string code_challenge_method)
        {
            string authorizationRequest = string.Format("{0}?response_type=code&scope=openid%20profile&redirect_uri={1}&client_id={2}&state={3}&code_challenge={4}&code_challenge_method={5}",
                 authorizationEndpoint,
                 Uri.EscapeDataString(redirectURI),
                 clientID,
                 state,
                 code_challenge,
                 code_challenge_method);

            return authorizationRequest;
        }

        public async Task<string> GetAccessToken()
        {
            // 1.Generates state and PKCE values.
            // More details https://developers.google.com/identity/protocols/oauth2/native-app
            string state = GmailUtils.randomDataBase64url(32);
            string code_verifier = GmailUtils.randomDataBase64url(32);
            string code_challenge = GmailUtils.base64urlencodeNoPadding(GmailUtils.sha256(code_verifier));
            const string code_challenge_method = "S256";

            // 2. Creates a redirect URI using an available port on the loopback address.
            string redirectURI = string.Format("http://{0}:{1}/", IPAddress.Loopback, GmailUtils.GetRandomUnusedPort());
            Output("redirect URI: " + redirectURI);

            // 3.Creates an HttpListener to listen for requests on that redirect URI.
            var http = new HttpListener();
            http.Prefixes.Add(redirectURI);
            Output("Listening..");
            http.Start();

            // 4. Creates the OAuth 2.0 authorization request.
            var authorizationRequest = CreateGmailOAth2AuthorizationRequest(redirectURI, state, code_challenge, code_challenge_method);

            var option = new ChromeOptions();
            option.DebuggerAddress = @"127.0.0.1:9222";
            //option.DebuggerAddress = "http://localhost:9222";
            WebDriver driver = new ChromeDriver(option);

            //driver.Navigate().GoToUrl("https://accounts.google.com/");
            driver.Navigate().GoToUrl(authorizationRequest);
            //Process.Start("\"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe\"", authorizationRequest);




            var context = await http.GetContextAsync();

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
            var accessToken = await performCodeExchange(code, code_verifier, redirectURI);

            return accessToken;
        }

        public string CreateGmailOAth2AuthorizationRequest(string redirectURI, string state, string code_challenge, string code_challenge_method)
        {
            //scope https://mail.google.com/
            //https://www.googleapis.com/auth/gmail.modify
            //https://www.googleapis.com/auth/gmail.readonly
            //https://www.googleapis.com/auth/gmail.metadata
            string authorizationRequest = string.Format("{0}?response_type=code&scope=https://mail.google.com/%20https://www.googleapis.com/auth/gmail.modify%20https://www.googleapis.com/auth/gmail.readonly%20https://www.googleapis.com/auth/gmail.metadata%20profile&redirect_uri={1}&client_id={2}&state={3}&code_challenge={4}&code_challenge_method={5}",
                 authorizationEndpoint,
                 Uri.EscapeDataString(redirectURI),
                 clientID,
                 state,
                 code_challenge,
                 code_challenge_method);

            return authorizationRequest;
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
                    userinfoCall(access_token);

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

        private async void userinfoCall(string access_token)
        {
            Output("Making API Call to Userinfo...");

            // builds the  request
            string userinfoRequestURI = "https://www.googleapis.com/oauth2/v3/userinfo";

            // sends the request
            HttpWebRequest userinfoRequest = (HttpWebRequest)WebRequest.Create(userinfoRequestURI);
            userinfoRequest.Method = "GET";
            userinfoRequest.Headers.Add(string.Format("Authorization: Bearer {0}", access_token));
            userinfoRequest.ContentType = "application/x-www-form-urlencoded";
            userinfoRequest.Accept = "Accept=text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";

            // gets the response
            WebResponse userinfoResponse = await userinfoRequest.GetResponseAsync();
            using (StreamReader userinfoResponseReader = new StreamReader(userinfoResponse.GetResponseStream()))
            {
                // reads response body
                string userinfoResponseText = await userinfoResponseReader.ReadToEndAsync();

                Output(userinfoResponseText);
            }
        }


        
        public async Task<string> GetGmailVerificationCode(string access_token)
        {
            try
            {
                Output("Making API Call to Userinfo...");

                // builds the  request
                string gmailRequestURI = "https://gmail.googleapis.com//gmail/v1/users/me/messages";

                // sends the request
                HttpWebRequest gmailRequest = (HttpWebRequest)WebRequest.Create(gmailRequestURI);
                gmailRequest.Method = "GET";
                gmailRequest.Headers.Add(string.Format("Authorization: Bearer {0}", access_token));
                gmailRequest.ContentType = "application/x-www-form-urlencoded";
                gmailRequest.Accept = "Accept=text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";

                Message firstMessage = null;
                // gets the response
                WebResponse userinfoResponse = await gmailRequest.GetResponseAsync();
                using StreamReader userinfoResponseReader = new StreamReader(userinfoResponse.GetResponseStream());
                // reads response body
                string gmailResponseText = await userinfoResponseReader.ReadToEndAsync();

                // TODO: TADA get list messages
                var myUnReadMsg = JsonConvert.DeserializeObject<ListMessagesResponse>(gmailResponseText);

                Output(gmailResponseText);

                firstMessage = myUnReadMsg?.Messages?.FirstOrDefault();


                var threadId = firstMessage?.ThreadId;
                var detailUrl = $"https://gmail.googleapis.com/gmail/v1/users/me/messages/{threadId}";

                using var client = new HttpClient();
                var response = await client.GetAsync(detailUrl);

                var resStr = await response.Content.ReadAsStringAsync();

                // TOD using 
                var hideMyAccBaseRes = JsonConvert.DeserializeObject<HideMyAccBaseRes<HideMyAccProfile>>(resStr);

                return gmailResponseText;


            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Appends the given string to the on-screen log, and the debug console.
        /// </summary>
        /// <param name="output">string to be appended</param>
        public void Output(string output)
        {
            //outputBox.Text = outputBox.Text + output + Environment.NewLine;
            Console.WriteLine(output);
        }

        public async Task<GmailServices> GetGmailServiceAsync()
        {
            UserCredential credential;
            var secretPath = "D:\\Wordspace\\_be\\AutoAppenWinform\\AutoAppenWinform\\ClientSecrets\\client_secret.json";

            //string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
            //credPath = Path.Combine(credPath, "ClientSecrets\\client_secret.json");
            using (var stream = new FileStream(secretPath, FileMode.Open, FileAccess.Read))
            {
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    new[] { GmailServices.Scope.GmailReadonly },
                    "user", CancellationToken.None, new FileDataStore("Gmail.ListMyLibrary"));
            }

            // Create the service.
            var gmailService = new GmailServices(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Auto Appen",
            });

            return gmailService;
        }


        // TODO more detail https://developers.google.com/gmail/api/reference/rest/v1/users.messages/list
        public void GetUnreadMessages()
        {
            var messages = _gmailServiceCon.Users.Messages.List("me");

            messages.LabelIds = "INBOX";
            messages.Q = "is:unread";

            var result = messages.Execute();

            if (result != null && result.Messages != null)
            {
                List<Message> messageRes = new List<Message>();
                // get gmail details
                foreach (var message in result.Messages)
                {
                    GetMessageDetail(message.Id);
                }
            }
        }

        public string ReadCodeFromMessage()
        {
            return string.Empty;
        }

        public void TakeScreenShot(WebDriver driver)
        {
            string subPath = @"C://Apppen/"; // Your code goes here

            string imgUrl = @"C://Apppen/appen.png";

            bool exists = Directory.Exists(subPath);

            if (!exists)
                Directory.CreateDirectory(subPath);

            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(imgUrl, ScreenshotImageFormat.Png);

            var text = ScanQRCode(imgUrl);
        }

        public string ScanQRCode(string imgUrl)
        {
            // more details: https://csharp.hotexamples.com/examples/ZXing/BinaryBitmap/-/php-binarybitmap-class-examples.html
            imgUrl = @"C://Apppen/vn.png";
            var img = Image.FromFile(imgUrl);
            var bmap = new Bitmap(img);
            var ms = new MemoryStream();
            bmap.Save(ms, ImageFormat.Bmp);
            var bytes = ms.GetBuffer();
            LuminanceSource source = new RGBLuminanceSource(bytes, bmap.Width, bmap.Height);
            var bitmap = new BinaryBitmap(new HybridBinarizer(source));
            var result = new MultiFormatReader().decode(bitmap);
            return result.Text;
        }

        public void GetMessageDetail(string messageId)
        {
            var emailInfoReq = _gmailServiceCon.Users.Messages.Get("me", messageId);
            var emailInfoResponse = emailInfoReq.Execute();

            if (emailInfoResponse != null)
            {
                string from = "";
                string date = "";
                string subject = "";
                string body = "";
                //loop through the headers and get the fields we need...
                foreach (var mParts in emailInfoResponse.Payload.Headers)
                {
                    if (mParts.Name == "Date")
                    {
                        date = mParts.Value;
                    }
                    else if (mParts.Name == "From")
                    {
                        from = mParts.Value;
                    }
                    else if (mParts.Name == "Subject")
                    {
                        subject = mParts.Value;
                    }

                    if (date != "" && from != "")
                    {
                        if (emailInfoResponse.Payload.Parts == null && emailInfoResponse.Payload.Body != null)
                            body = GmailUtils.DecodeBase64String(emailInfoResponse.Payload.Body.Data);
                        else
                            body = GetNestedBodyParts(emailInfoResponse.Payload.Parts, "");

                        //now you have the data you want....
                    }
                }
            }
        }

        public string GetNestedBodyParts(IList<MessagePart> part, string curr)
        {
            string str = curr;
            if (part == null)
                return str;

            foreach (var parts in part)
            {
                if (parts.Parts == null)
                {
                    if (parts.Body != null && parts.Body.Data != null)
                    {
                        var ts = GmailUtils.DecodeBase64String(parts.Body.Data);
                        str += ts;
                    }
                }
                else
                {
                    return GetNestedBodyParts(parts.Parts, str);
                }
            }

            return str;
        }

        public void SayHello()
        {
            Console.WriteLine("Hello");
        }
    }
}