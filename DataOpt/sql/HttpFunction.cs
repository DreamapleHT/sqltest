using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Net;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Diagnostics;
namespace DataOpt.sql
{
    class HttpFunction
    {
        #region 解析http请求

        private static readonly string DefaultUserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36";
        /// <summary>
        /// 测试网络是否可用
        /// </summary>
        /// <returns></returns>
        public static bool TestNetworkStatus()
        {
            var urls = new string[] { "http://www.baidu.com", "http://www.qq.com/" };
            int count = 0;
            foreach (var url in urls)
            {
                try
                {
                    var resp = HttpResponseGet(url, method: "HEAD");
                    if (resp.StatusCode == HttpStatusCode.OK)
                    {
                        count++;
                    }
                }
                catch { }
            }
            return count > 0;
        }

        /// <summary>  
        /// 创建GET方式的HTTP请求  
        /// </summary>  
        /// <param name="url">请求的URL</param>  
        /// <param name="timeout">请求的超时时间</param>  
        /// <param name="userAgent">请求的客户端浏览器信息，可以为空</param>  
        /// <param name="cookies">随同HTTP请求发送的Cookie信息，如果不需要身份验证可以为空</param>  
        /// <returns></returns>  
        public static string HttpResponseGetString(string url, string getData = null, int? timeout = null, string userAgent = null,
            CookieCollection cookies = null, string method = "GET")
        {
            using (var response = HttpResponseGet(url, getData, timeout, userAgent, cookies, method))
            {
                return GetResponseString(response);
            }
        }

        /// <summary>  
        /// 创建GET方式的HTTP请求  
        /// </summary>  
        /// <param name="url">请求的URL</param>  
        /// <param name="timeout">请求的超时时间</param>  
        /// <param name="userAgent">请求的客户端浏览器信息，可以为空</param>  
        /// <param name="cookies">随同HTTP请求发送的Cookie信息，如果不需要身份验证可以为空</param>  
        /// <returns></returns>  
        public static HttpWebResponse HttpResponseGet(string url, string getData = null, int? timeout = null, string userAgent = null,
            CookieCollection cookies = null, string method = "GET")
        {
            if (string.IsNullOrEmpty(url))
            {
                throw new ArgumentNullException("url");
            }
            if (getData != null)
            {
                url = url.Contains('?') ? url.Trim('&') : url + "?";
                url += getData.Trim('&');
            }
            HttpWebRequest request = null;
            //如果是发送HTTPS请求  
            if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
            {
                ServicePointManager.ServerCertificateValidationCallback =
                    new RemoteCertificateValidationCallback(CheckValidationResult);
                request = WebRequest.Create(url) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version10;
            }
            else
            {
                request = WebRequest.Create(url) as HttpWebRequest;
            }
            request.Method = method;
            if (!string.IsNullOrEmpty(userAgent))
            {
                request.UserAgent = userAgent;
            }
            else
            {
                request.UserAgent = DefaultUserAgent;
            }
            if (timeout.HasValue)
            {
                request.Timeout = timeout.Value;
            }
            if (cookies != null)
            {
                request.CookieContainer = new CookieContainer();
                request.CookieContainer.Add(cookies);
            }
            return request.GetResponse() as HttpWebResponse;
        }

        /// <summary>  
        /// 创建POST方式的HTTP请求  
        /// </summary>  
        /// <param name="url">请求的URL</param>  
        /// <param name="parameters">随同请求POST的参数名称及参数值字典</param>  
        /// <param name="timeout">请求的超时时间</param>  
        /// <param name="userAgent">请求的客户端浏览器信息，可以为空</param>  
        /// <param name="requestEncoding">发送HTTP请求时所用的编码</param>  
        /// <param name="cookies">随同HTTP请求发送的Cookie信息，如果不需要身份验证可以为空</param>  
        /// <returns></returns>  
        public static string HttpResponsePostString(string url, string postData = null,
            Encoding requestEncoding = null, int? timeout = null, string userAgent = null,
            CookieCollection cookies = null, bool sendJson = false, bool sendXml = false)
        {
            using (var response = HttpResponsePost(url, postData, requestEncoding, timeout, userAgent, cookies, sendJson, sendXml))
            {
                return GetResponseString(response);
            }
        }

        /// <summary>  
        /// 创建POST方式的HTTP请求  
        /// </summary>  
        /// <param name="url">请求的URL</param>  
        /// <param name="parameters">随同请求POST的参数名称及参数值字典</param>  
        /// <param name="timeout">请求的超时时间</param>  
        /// <param name="userAgent">请求的客户端浏览器信息，可以为空</param>  
        /// <param name="requestEncoding">发送HTTP请求时所用的编码</param>  
        /// <param name="cookies">随同HTTP请求发送的Cookie信息，如果不需要身份验证可以为空</param>  
        /// <returns></returns>  
        public static HttpWebResponse HttpResponsePost(string url, string postData = null,
            Encoding requestEncoding = null, int? timeout = null, string userAgent = null,
            CookieCollection cookies = null, bool sendJson = false, bool sendXml = false)
        {
            if (string.IsNullOrEmpty(url))
            {
                throw new ArgumentNullException("url");
            }
            if (requestEncoding == null)
            {
                requestEncoding = Encoding.UTF8;
            }
            HttpWebRequest request = null;
            //如果是发送HTTPS请求  
            if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
            {
                ServicePointManager.ServerCertificateValidationCallback =
                    new RemoteCertificateValidationCallback(CheckValidationResult);
                request = WebRequest.Create(url) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version10;
            }
            else
            {
                request = WebRequest.Create(url) as HttpWebRequest;
            }
            request.Method = "POST";
            if (sendJson)
            {
                request.ContentType = "application/json";
            }
            else if (sendXml)
            {
                request.ContentType = "application/xml";
            }
            else
            {
                request.ContentType = "application/x-www-form-urlencoded";
            }
            if (!string.IsNullOrEmpty(userAgent))
            {
                request.UserAgent = userAgent;
            }
            else
            {
                request.UserAgent = DefaultUserAgent;
            }

            if (timeout.HasValue)
            {
                request.Timeout = timeout.Value;
            }
            if (cookies != null)
            {
                request.CookieContainer = new CookieContainer();
                request.CookieContainer.Add(cookies);
            }
            //如果需要POST数据  
            if (!string.IsNullOrEmpty(postData))
            {
                byte[] data = requestEncoding.GetBytes(postData);
                using (Stream stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }
            }
            return request.GetResponse() as HttpWebResponse;
        }

        /// <summary>
        /// 从Response中获取字符串结果
        /// </summary>
        /// <param name="response"></param>
        /// <returns></returns>
        private static string GetResponseString(HttpWebResponse response)
        {
            Stream stream = response.GetResponseStream();
            if (stream != null)
            {
                StreamReader sr = new StreamReader(stream);
                string result = sr.ReadToEnd();

                stream.Close();
                stream.Dispose();
                sr.Close();
                sr.Dispose();

                return result;
            }
            return string.Empty;
        }

        /// <summary>
        /// 屏蔽https的服务器证书验证,总是返回true
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="certificate"></param>
        /// <param name="chain"></param>
        /// <param name="errors"></param>
        /// <returns></returns>
        private static bool CheckValidationResult(object sender, X509Certificate certificate, X509Chain chain,
            SslPolicyErrors errors)
        {
            return true; //总是接受  
        }

        #endregion
    }
}
