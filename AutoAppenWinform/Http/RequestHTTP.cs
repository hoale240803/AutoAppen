using System.Diagnostics;
using System.Net.Cache;
using System.Net.Security;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Text;

namespace AutoAppenWinform.Http
{
    public class RequestHTTP
    {
        private CookieCollection RequestCookie = new CookieCollection();

        private CookieCollection ResponseCookie = new CookieCollection();

        private WebHeaderCollection RequestHeaders = new WebHeaderCollection();

        private WebHeaderCollection ResponseHeaders = new WebHeaderCollection();

        private CookieContainer Cookies = new CookieContainer();

        private string CookieString = null;

        private WebProxy wProxy;

        public HttpStatusCode code;

        public bool usProxy = false;

        private string[] DefaultHeaders;

        private bool keepAlive = false;

        private NetworkCredential networkCredential = null;

        public RequestHTTP()
        {
            ServicePointManager.DefaultConnectionLimit = 1000;
            ServicePointManager.Expect100Continue = false;
        }

        public static void IgnoreBadCertificates()
        {
            ServicePointManager.ServerCertificateValidationCallback = AcceptAllCertifications;
        }

        private static bool AcceptAllCertifications(object sender, X509Certificate certification, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        public void SetSSL(SecurityProtocolType type)
        {
            ServicePointManager.SecurityProtocol = type;
        }

        public void SetKeepAlive(bool k)
        {
            keepAlive = k;
        }

        public void SetCredential(NetworkCredential credential)
        {
            networkCredential = credential;
        }

        public async Task<string> RequestAsync(string Method = "GET", string Url = "", string[] Headers = null, byte[] Data = null, bool autoredrirect = true, WebProxy proxy = null, int time_out = 60000)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            IgnoreBadCertificates();
            if (Url == null)
            {
                return "";
            }

            HttpWebRequest hWRQ;
            try
            {
                hWRQ = (HttpWebRequest)WebRequest.Create(Url);
                hWRQ.AllowAutoRedirect = autoredrirect;
                hWRQ.CookieContainer = Cookies;
                hWRQ.KeepAlive = keepAlive;
                hWRQ.ReadWriteTimeout = time_out;
                hWRQ.ContinueTimeout = time_out;
                hWRQ.Timeout = time_out;
                if (networkCredential != null && networkCredential != null)
                {
                    string encoded = Convert.ToBase64String(Encoding.GetEncoding("ISO-8859-1").GetBytes(networkCredential.UserName + ":" + networkCredential.Password));
                    hWRQ.Headers.Add("Authorization", "Basic " + encoded);
                }

                if (proxy != null)
                {
                    hWRQ.Proxy = proxy;
                }
                else if (usProxy && wProxy != null)
                {
                    hWRQ.Proxy = wProxy;
                }

                hWRQ.Method = Method;
                HttpRequestCachePolicy policy = (HttpRequestCachePolicy)(HttpWebRequest.DefaultCachePolicy = new HttpRequestCachePolicy(HttpRequestCacheLevel.Default));
                HttpRequestCachePolicy noCachePolicy = (HttpRequestCachePolicy)(hWRQ.CachePolicy = new HttpRequestCachePolicy(HttpRequestCacheLevel.NoCacheNoStore));
                hWRQ.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                if (DefaultHeaders != null && DefaultHeaders.Length != 0)
                {
                    for (int j = 0; j < DefaultHeaders.Length; j++)
                    {
                        SetRequestHeaders(hWRQ, DefaultHeaders[j]);
                    }
                }

                if (Headers != null && Headers.Length != 0)
                {
                    for (int i = 0; i < Headers.Length; i++)
                    {
                        SetRequestHeaders(hWRQ, Headers[i]);
                    }
                }

                if (RequestCookie.Count > 0)
                {
                    hWRQ.CookieContainer.Add(RequestCookie);
                }

                if (Data != null)
                {
                    if (Headers == null)
                    {
                        SetRequestHeaders(hWRQ, "Content-Type: application/x-www-form-urlencoded; charset=UTF-8");
                    }

                    hWRQ.ContentLength = Data.Length;
                    hWRQ.GetRequestStream().Write(Data, 0, Data.Length);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return "!Exception!" + ex.Message;
            }

            ServicePointManager.ServerCertificateValidationCallback = (RemoteCertificateValidationCallback)Delegate.Combine(ServicePointManager.ServerCertificateValidationCallback, (RemoteCertificateValidationCallback)((sender, cert, chain, sslPolicyErrors) => true));
            HttpWebResponse hWRP = null;
            string hWRP_Text = "";
            try
            {
                hWRP = (HttpWebResponse)await hWRQ.GetResponseAsync();
                code = hWRP.StatusCode;
            }
            catch (WebException e)
            {
                try
                {
                    hWRP_Text = new StreamReader(e.Response.GetResponseStream()).ReadToEnd();
                    ResponseHeaders = e.Response.Headers;
                }
                catch (Exception)
                {
                }

                if (e.Status == WebExceptionStatus.ProtocolError)
                {
                    HttpWebResponse response = e.Response as HttpWebResponse;
                    if (response != null)
                    {
                        code = response.StatusCode;
                    }
                    else
                    {
                        code = HttpStatusCode.NotFound;
                    }
                }

                hWRQ.Abort();
                return hWRP_Text;
            }
            catch (Exception)
            {
            }

            try
            {
                hWRP_Text = GetResponseText(hWRP);
                RequestHeaders = hWRQ.Headers;
                ResponseHeaders = hWRP.Headers;
                ResponseCookie = hWRP.Cookies;
                Cookies = hWRQ.CookieContainer;
                CookieString = GetCookiesString(hWRQ.RequestUri.ToString());
            }
            catch (Exception)
            {
            }

            hWRQ.Abort();
            hWRP.Dispose();
            return hWRP_Text;
        }

        public void SetCookieCapacity(int cap)
        {
            Cookies.Capacity = cap;
        }

        public void SetCookieMaxLength(int length)
        {
            Cookies.MaxCookieSize = length;
        }

        public void SetCookiePerDomainCapcity(int cap)
        {
            Cookies.PerDomainCapacity = cap;
        }

        public string Request(string Method = "GET", string Url = "", string[] Headers = null, byte[] Data = null, bool autoredrirect = true, WebProxy proxy = null, int time_out = 60000)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            IgnoreBadCertificates();
            if (Url == null)
            {
                return "";
            }

            HttpWebRequest httpWebRequest;
            try
            {
                httpWebRequest = (HttpWebRequest)WebRequest.Create(Url);
                httpWebRequest.AllowAutoRedirect = autoredrirect;
                httpWebRequest.CookieContainer = Cookies;
                httpWebRequest.KeepAlive = keepAlive;
                httpWebRequest.ReadWriteTimeout = time_out;
                httpWebRequest.ContinueTimeout = time_out;
                httpWebRequest.Timeout = time_out;
                if (networkCredential != null)
                {
                    string text = Convert.ToBase64String(Encoding.GetEncoding("ISO-8859-1").GetBytes(networkCredential.UserName + ":" + networkCredential.Password));
                    httpWebRequest.Headers.Add("Authorization", "Basic " + text);
                }

                if (proxy != null)
                {
                    httpWebRequest.Proxy = proxy;
                }
                else if (usProxy && wProxy != null)
                {
                    httpWebRequest.Proxy = wProxy;
                }

                httpWebRequest.Method = Method;
                HttpRequestCachePolicy httpRequestCachePolicy = (HttpRequestCachePolicy)(HttpWebRequest.DefaultCachePolicy = new HttpRequestCachePolicy(HttpRequestCacheLevel.Default));
                HttpRequestCachePolicy httpRequestCachePolicy2 = (HttpRequestCachePolicy)(httpWebRequest.CachePolicy = new HttpRequestCachePolicy(HttpRequestCacheLevel.NoCacheNoStore));
                httpWebRequest.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                if (DefaultHeaders != null && DefaultHeaders.Length != 0)
                {
                    for (int i = 0; i < DefaultHeaders.Length; i++)
                    {
                        SetRequestHeaders(httpWebRequest, DefaultHeaders[i]);
                    }
                }

                if (Headers != null && Headers.Length != 0)
                {
                    for (int j = 0; j < Headers.Length; j++)
                    {
                        SetRequestHeaders(httpWebRequest, Headers[j]);
                    }
                }

                if (RequestCookie.Count > 0)
                {
                    httpWebRequest.CookieContainer.Add(RequestCookie);
                }

                if (Data != null)
                {
                    if (Headers == null)
                    {
                        SetRequestHeaders(httpWebRequest, "Content-Type: application/x-www-form-urlencoded; charset=UTF-8");
                    }

                    httpWebRequest.ContentLength = Data.Length;
                    Stream requestStream = httpWebRequest.GetRequestStream();
                    requestStream.Write(Data, 0, Data.Length);
                    requestStream.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                return "!Exception!" + ex.Message;
            }

            ServicePointManager.ServerCertificateValidationCallback = (RemoteCertificateValidationCallback)Delegate.Combine(ServicePointManager.ServerCertificateValidationCallback, (RemoteCertificateValidationCallback)((sender, cert, chain, sslPolicyErrors) => true));
            HttpWebResponse httpWebResponse = null;
            string result = "";
            try
            {
                httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                code = httpWebResponse.StatusCode;
            }
            catch (WebException ex2)
            {
                try
                {
                    result = new StreamReader(ex2.Response.GetResponseStream()).ReadToEnd();
                    ResponseHeaders = ex2.Response.Headers;
                }
                catch (Exception)
                {
                }

                if (ex2.Status == WebExceptionStatus.ProtocolError)
                {
                    HttpWebResponse httpWebResponse2 = ex2.Response as HttpWebResponse;
                    if (httpWebResponse2 != null)
                    {
                        code = httpWebResponse2.StatusCode;
                    }
                    else
                    {
                        code = HttpStatusCode.NotFound;
                    }
                }

                httpWebRequest.Abort();
                return result;
            }
            catch (Exception ex4)
            {
                Console.WriteLine(ex4.Message);
                Console.WriteLine(ex4.StackTrace);
            }

            try
            {
                result = GetResponseText(httpWebResponse);
                RequestHeaders = httpWebRequest.Headers;
                ResponseHeaders = httpWebResponse.Headers;
                ResponseCookie = httpWebResponse.Cookies;
                Cookies = httpWebRequest.CookieContainer;
                CookieString = GetCookiesString(httpWebRequest.RequestUri.ToString());
            }
            catch (Exception)
            {
            }

            try
            {
                httpWebRequest.Abort();
                httpWebResponse.Dispose();
            }
            catch (Exception)
            {
            }

            return result;
        }

        public async Task<byte[]> Request_byteAsync(string Method = "GET", string Url = "", string[] Headers = null, byte[] Data = null, bool autoredrirect = true, WebProxy proxy = null, int time_out = 60000)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            IgnoreBadCertificates();
            if (Url == null)
            {
                return null;
            }

            HttpWebRequest hWRQ;
            try
            {
                hWRQ = (HttpWebRequest)WebRequest.Create(Url);
                hWRQ.AllowAutoRedirect = autoredrirect;
                hWRQ.CookieContainer = Cookies;
                hWRQ.KeepAlive = keepAlive;
                hWRQ.ReadWriteTimeout = time_out;
                hWRQ.ContinueTimeout = time_out;
                hWRQ.Timeout = time_out;
                if (networkCredential != null && networkCredential != null)
                {
                    string encoded = Convert.ToBase64String(Encoding.GetEncoding("ISO-8859-1").GetBytes(networkCredential.UserName + ":" + networkCredential.Password));
                    hWRQ.Headers.Add("Authorization", "Basic " + encoded);
                }

                if (proxy != null)
                {
                    hWRQ.Proxy = proxy;
                }
                else if (usProxy && wProxy != null)
                {
                    hWRQ.Proxy = wProxy;
                }

                hWRQ.Method = Method;
                HttpRequestCachePolicy policy = (HttpRequestCachePolicy)(HttpWebRequest.DefaultCachePolicy = new HttpRequestCachePolicy(HttpRequestCacheLevel.Default));
                HttpRequestCachePolicy noCachePolicy = (HttpRequestCachePolicy)(hWRQ.CachePolicy = new HttpRequestCachePolicy(HttpRequestCacheLevel.NoCacheNoStore));
                hWRQ.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                if (DefaultHeaders != null && DefaultHeaders.Length != 0)
                {
                    for (int j = 0; j < DefaultHeaders.Length; j++)
                    {
                        SetRequestHeaders(hWRQ, DefaultHeaders[j]);
                    }
                }

                if (Headers != null && Headers.Length != 0)
                {
                    for (int i = 0; i < Headers.Length; i++)
                    {
                        SetRequestHeaders(hWRQ, Headers[i]);
                    }
                }

                if (RequestCookie.Count > 0)
                {
                    hWRQ.CookieContainer.Add(RequestCookie);
                }

                if (Data != null)
                {
                    if (Headers == null)
                    {
                        SetRequestHeaders(hWRQ, "Content-Type: application/x-www-form-urlencoded; charset=UTF-8");
                    }

                    hWRQ.ContentLength = Data.Length;
                    hWRQ.GetRequestStream().Write(Data, 0, Data.Length);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return Encoding.ASCII.GetBytes("!Exception!" + ex.Message);
            }

            ServicePointManager.ServerCertificateValidationCallback = (RemoteCertificateValidationCallback)Delegate.Combine(ServicePointManager.ServerCertificateValidationCallback, (RemoteCertificateValidationCallback)((sender, cert, chain, sslPolicyErrors) => true));
            HttpWebResponse hWRP = null;
            byte[] hWRP_Text = null;
            try
            {
                hWRP = (HttpWebResponse)await hWRQ.GetResponseAsync();
                code = hWRP.StatusCode;
            }
            catch (WebException e)
            {
                try
                {
                    MemoryStream ms = new MemoryStream();
                    Task task = e.Response.GetResponseStream().CopyToAsync(ms);
                    Stopwatch sw = new Stopwatch();
                    sw.Start();
                    while (true)
                    {
                        if (sw.ElapsedMilliseconds > time_out)
                        {
                            task.Dispose();
                            break;
                        }

                        if (task.IsCanceled || task.IsCompleted || task.IsFaulted)
                        {
                            task.Dispose();
                            break;
                        }
                    }

                    sw.Stop();
                    hWRP_Text = ms.ToArray();
                    ms.Dispose();
                    ResponseHeaders = e.Response.Headers;
                }
                catch (Exception)
                {
                }

                if (e.Status == WebExceptionStatus.ProtocolError)
                {
                    HttpWebResponse response = e.Response as HttpWebResponse;
                    if (response != null)
                    {
                        code = response.StatusCode;
                    }
                    else
                    {
                        code = HttpStatusCode.NotFound;
                    }
                }

                hWRQ.Abort();
                return hWRP_Text;
            }
            catch (Exception)
            {
            }

            try
            {
                hWRP_Text = GetResponseByte(hWRP, time_out);
                RequestHeaders = hWRQ.Headers;
                ResponseHeaders = hWRP.Headers;
                ResponseCookie = hWRP.Cookies;
                Cookies = hWRQ.CookieContainer;
                CookieString = GetCookiesString(hWRQ.RequestUri.ToString());
            }
            catch (Exception)
            {
            }

            hWRQ.Abort();
            return hWRP_Text;
        }

        public byte[] Request_byte(string Method = "GET", string Url = "", string[] Headers = null, byte[] Data = null, bool autoredrirect = true, WebProxy proxy = null, int time_out = 60000)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            IgnoreBadCertificates();
            if (Url == null)
            {
                return null;
            }

            HttpWebRequest httpWebRequest;
            try
            {
                httpWebRequest = (HttpWebRequest)WebRequest.Create(Url);
                httpWebRequest.AllowAutoRedirect = autoredrirect;
                httpWebRequest.CookieContainer = Cookies;
                httpWebRequest.KeepAlive = keepAlive;
                httpWebRequest.ReadWriteTimeout = time_out;
                httpWebRequest.ContinueTimeout = time_out;
                httpWebRequest.Timeout = time_out;
                if (networkCredential != null && networkCredential != null)
                {
                    string text = Convert.ToBase64String(Encoding.GetEncoding("ISO-8859-1").GetBytes(networkCredential.UserName + ":" + networkCredential.Password));
                    httpWebRequest.Headers.Add("Authorization", "Basic " + text);
                }

                if (proxy != null)
                {
                    httpWebRequest.Proxy = proxy;
                }
                else if (usProxy && wProxy != null)
                {
                    httpWebRequest.Proxy = wProxy;
                }

                httpWebRequest.Method = Method;
                HttpRequestCachePolicy httpRequestCachePolicy = (HttpRequestCachePolicy)(HttpWebRequest.DefaultCachePolicy = new HttpRequestCachePolicy(HttpRequestCacheLevel.Default));
                HttpRequestCachePolicy httpRequestCachePolicy2 = (HttpRequestCachePolicy)(httpWebRequest.CachePolicy = new HttpRequestCachePolicy(HttpRequestCacheLevel.NoCacheNoStore));
                httpWebRequest.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                if (DefaultHeaders != null && DefaultHeaders.Length != 0)
                {
                    for (int i = 0; i < DefaultHeaders.Length; i++)
                    {
                        SetRequestHeaders(httpWebRequest, DefaultHeaders[i]);
                    }
                }

                if (Headers != null && Headers.Length != 0)
                {
                    for (int j = 0; j < Headers.Length; j++)
                    {
                        SetRequestHeaders(httpWebRequest, Headers[j]);
                    }
                }

                if (RequestCookie.Count > 0)
                {
                    httpWebRequest.CookieContainer.Add(RequestCookie);
                }

                if (Data != null)
                {
                    if (Headers == null)
                    {
                        SetRequestHeaders(httpWebRequest, "Content-Type: application/x-www-form-urlencoded; charset=UTF-8");
                    }

                    httpWebRequest.ContentLength = Data.Length;
                    httpWebRequest.GetRequestStream().Write(Data, 0, Data.Length);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return Encoding.ASCII.GetBytes("!Exception!" + ex.Message);
            }

            ServicePointManager.ServerCertificateValidationCallback = (RemoteCertificateValidationCallback)Delegate.Combine(ServicePointManager.ServerCertificateValidationCallback, (RemoteCertificateValidationCallback)((sender, cert, chain, sslPolicyErrors) => true));
            HttpWebResponse httpWebResponse = null;
            byte[] result = null;
            try
            {
                httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                code = httpWebResponse.StatusCode;
            }
            catch (WebException ex2)
            {
                try
                {
                    MemoryStream memoryStream = new MemoryStream();
                    Task task = ex2.Response.GetResponseStream().CopyToAsync(memoryStream);
                    Stopwatch stopwatch = new Stopwatch();
                    stopwatch.Start();
                    while (true)
                    {
                        if (stopwatch.ElapsedMilliseconds > time_out)
                        {
                            task.Dispose();
                            break;
                        }

                        if (task.IsCanceled || task.IsCompleted || task.IsFaulted)
                        {
                            task.Dispose();
                            break;
                        }
                    }

                    stopwatch.Stop();
                    stopwatch = null;
                    result = memoryStream.ToArray();
                    memoryStream.Dispose();
                    ResponseHeaders = ex2.Response.Headers;
                }
                catch (Exception)
                {
                }

                if (ex2.Status == WebExceptionStatus.ProtocolError)
                {
                    HttpWebResponse httpWebResponse2 = ex2.Response as HttpWebResponse;
                    if (httpWebResponse2 != null)
                    {
                        code = httpWebResponse2.StatusCode;
                    }
                    else
                    {
                        code = HttpStatusCode.NotFound;
                    }
                }

                httpWebRequest.Abort();
                return result;
            }
            catch (Exception)
            {
            }

            try
            {
                result = GetResponseByte(httpWebResponse, time_out);
                RequestHeaders = httpWebRequest.Headers;
                ResponseHeaders = httpWebResponse.Headers;
                ResponseCookie = httpWebResponse.Cookies;
                Cookies = httpWebRequest.CookieContainer;
                CookieString = GetCookiesString(httpWebRequest.RequestUri.ToString());
            }
            catch (Exception)
            {
            }

            httpWebRequest.Abort();
            return result;
        }

        public void ClearCookie()
        {
            Cookies = new CookieContainer();
            CookieString = null;
        }

        public void RemoveCookie(string url)
        {
            CookieCollection cookies = Cookies.GetCookies(new Uri(url));
            foreach (Cookie item in cookies)
            {
                item.Expires = DateTime.Now.Subtract(TimeSpan.FromDays(1.0));
            }
        }

        public bool SetProxy(string ip, int port)
        {
            try
            {
                wProxy = new WebProxy(ip, port);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public CookieContainer GetCookiesContainer()
        {
            return Cookies;
        }

        public void SetDefaultHeaders(string[] default_headers)
        {
            DefaultHeaders = default_headers;
        }

        public void SetCookiesContainer(CookieContainer CookContainer)
        {
            Cookies = CookContainer;
        }

        public CookieCollection GetResponseCookies()
        {
            return ResponseCookie;
        }

        public CookieCollection GetRequestCookies()
        {
            return RequestCookie;
        }

        public WebHeaderCollection GetRequestHeaders()
        {
            return RequestHeaders;
        }

        public WebHeaderCollection GetResponseHeaders()
        {
            return ResponseHeaders;
        }

        public void AddResponseCookies(CookieCollection Cookies)
        {
            ResponseCookie.Add(Cookies);
        }

        public void AddRequestCookies(CookieCollection addCookies)
        {
            RequestCookie.Add(addCookies);
        }

        private byte[] GetResponseByte(HttpWebResponse hWRP, int time_out)
        {
            byte[] result;
            try
            {
                MemoryStream memoryStream = new MemoryStream();
                Task task = hWRP.GetResponseStream().CopyToAsync(memoryStream);
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                while (true)
                {
                    if (stopwatch.ElapsedMilliseconds > time_out)
                    {
                        task.Dispose();
                        break;
                    }

                    if (task.IsCanceled || task.IsCompleted || task.IsFaulted)
                    {
                        task.Dispose();
                        break;
                    }
                }

                stopwatch.Stop();
                stopwatch = null;
                result = memoryStream.ToArray();
                memoryStream.Dispose();
            }
            catch (Exception ex)
            {
                result = Encoding.ASCII.GetBytes(ex.Message);
            }

            return result;
        }

        private string GetResponseText(HttpWebResponse hWRP)
        {
            string text = "";
            try
            {
                StreamReader streamReader = new StreamReader(hWRP.GetResponseStream());
                text = streamReader.ReadToEnd();
                streamReader.Close();
            }
            catch (Exception ex)
            {
                text = ex.Message;
            }

            return text;
        }

        public void SetRequestHeaders(HttpWebRequest wRq, string Header)
        {
            string[] array = Header.Split(':');
            if (array.Length <= 1)
            {
                return;
            }

            string text = "";
            for (int i = 1; i < array.Length; i++)
            {
                text = text + array[i] + ":";
            }

            text = text.Remove(text.Length - 1);
            if (text.StartsWith(" "))
            {
                text = text.Substring(1, text.Length - 1);
            }

            switch (array[0].ToLower())
            {
                case "accept":
                    wRq.Accept = text;
                    break;
                case "connection":
                    try
                    {
                        wRq.Connection = text;
                    }
                    catch (Exception)
                    {
                    }

                    break;
                case "content-length":
                    wRq.ContentLength = int.Parse(text);
                    break;
                case "content-type":
                    wRq.ContentType = text;
                    break;
                case "date":
                    {
                        DateTime dateTime4 = wRq.Date = Convert.ToDateTime(text);
                        break;
                    }
                case "expect":
                    wRq.Expect = text;
                    break;
                case "host":
                    wRq.Host = text;
                    break;
                case "if-modified-since":
                    {
                        DateTime dateTime2 = wRq.IfModifiedSince = Convert.ToDateTime(text);
                        break;
                    }
                case "range":
                    wRq.AddRange(int.Parse(text));
                    break;
                case "referer":
                    wRq.Referer = text;
                    break;
                case "transfer-encoding":
                    wRq.TransferEncoding = text;
                    break;
                case "user-agent":
                    wRq.UserAgent = text;
                    break;
                case "cookie":
                    SetCookie(text, Regex.Match(wRq.Host, "^(?:\\w+://)?([^/?]*)").Groups[1].Value.Replace("www.", ""));
                    wRq.CookieContainer = Cookies;
                    break;
                case "authorization":
                    wRq.Headers["authorization"] = text;
                    break;
                default:
                    wRq.Headers.Add(Header);
                    break;
            }
        }

        public string GetCookiesString()
        {
            return CookieString;
        }

        public string GetCookiesString(string url)
        {
            return Cookies.GetCookieHeader(new Uri(url));
        }

        public void SetCookie(string CookiesString, string domain)
        {
            CookieContainer cookieContainer = new CookieContainer();
            CookiesString = CookiesString.Trim();
            if (CookiesString.Last() != ';')
            {
                CookiesString += ";";
            }

            Regex regex = new Regex("(.*?)=(.*?);");
            MatchCollection matchCollection = regex.Matches(CookiesString);
            for (int i = 0; i < matchCollection.Count; i++)
            {
                GroupCollection groups = matchCollection[i].Groups;
                cookieContainer.Add(new Cookie(groups[1].Value.Trim(), groups[2].Value.Trim().Replace(",", "%2C"), "/", domain));
            }

            Cookies = cookieContainer;
        }

        public CookieCollection CookiesFromString(string CookiesString)
        {
            CookieCollection cookieCollection = new CookieCollection();
            CookiesString = CookiesString.Trim();
            if (CookiesString.Last() != ';')
            {
                CookiesString += ";";
            }

            Regex regex = new Regex("(.*?)=(.*?);");
            MatchCollection matchCollection = regex.Matches(CookiesString);
            for (int i = 0; i < matchCollection.Count; i++)
            {
                GroupCollection groups = matchCollection[i].Groups;
                Cookie cookie = new Cookie();
                cookie.Name = groups[1].Value.Trim();
                cookie.Value = groups[2].Value.Trim().Replace(",", "%2C");
                cookieCollection.Add(cookie);
            }

            return cookieCollection;
        }
    }
}