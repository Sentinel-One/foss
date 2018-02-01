using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net;
using System.Net.Security;

public enum HttpVerb
{
    GET,
    POST,
    PUT,
    DELETE
}

namespace S1ExcelPlugIn
{
  public class RestClientInterface
  {
    public string EndPoint { get; set; }
    public HttpVerb Method { get; set; }
    public string ContentType { get; set; }
    public string PostData { get; set; }

    public RestClientInterface()
    {
      EndPoint = "";
      Method = HttpVerb.GET;
      ContentType = "text/xml";
      PostData = "";
    }
    public RestClientInterface(string endpoint)
    {
      EndPoint = endpoint;
      Method = HttpVerb.GET;
      ContentType = "text/xml";
      PostData = "";
    }
    public RestClientInterface(string endpoint, HttpVerb method)
    {
      EndPoint = endpoint;
      Method = method;
      ContentType = "text/xml";
      PostData = "";
    }

    public RestClientInterface(string endpoint, HttpVerb method, string postData)
    {
      EndPoint = endpoint;
      Method = method;
        // ContentType = "text/xml";
        ContentType = "application/json";
        PostData = postData;
    }


    public string MakeRequest()
    {
      return MakeRequest("");
    }

    public string MakeRequest(string parameters)
    {
      // ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
      ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(delegate { return true; });

            System.Net.ServicePointManager.SecurityProtocol =
                System.Net.SecurityProtocolType.Tls |
                System.Net.SecurityProtocolType.Tls11 |
                System.Net.SecurityProtocolType.Tls12 |
                System.Net.SecurityProtocolType.Ssl3;

            var request = (HttpWebRequest)WebRequest.Create(EndPoint);

      request.Method = Method.ToString();
      // request.Headers["Authorization"] = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(parameters));
      request.Headers["Authorization"] = "Token " + parameters;
      request.ContentLength = 0;
      request.ContentType = ContentType;

      if (!string.IsNullOrEmpty(PostData) && Method == HttpVerb.POST)
      {
        var encoding = new UTF8Encoding();
        var bytes = Encoding.GetEncoding("iso-8859-1").GetBytes(PostData);
        request.ContentLength = bytes.Length;

        using (var writeStream = request.GetRequestStream())
        {
          writeStream.Write(bytes, 0, bytes.Length);
        }
      }

            try
            {
                using (var response = (HttpWebResponse)request.GetResponse())
                {
                    var responseValue = string.Empty;

                    if (response.StatusCode != HttpStatusCode.OK)
                    {
                        var message = String.Format("Request failed. Received HTTP {0}", response.StatusCode);
                        throw new ApplicationException(message + " - " + responseValue);
                    }

                    // grab the response
                    using (var responseStream = response.GetResponseStream())
                    {
                        if (responseStream != null)
                            using (var reader = new StreamReader(responseStream))
                            {
                                responseValue = reader.ReadToEnd();
                            }
                    }

                    return responseValue;
                }
            }
            catch (WebException e)
            {
                using (WebResponse response = e.Response)
                {
                    HttpWebResponse httpResponse = (HttpWebResponse)response;

                    if (httpResponse == null)
                    {
                        throw new ApplicationException("No response - check for valid server name/IP and port number");
                    }

                    using (Stream data = response.GetResponseStream())
                    using (var reader = new StreamReader(data))
                    {
                        string text = reader.ReadToEnd();
                        if (text.Contains("(404)"))
                        {
                            throw new ApplicationException("You have reached some server, but not the correct BigFix Server as it returned a 404 Page Not Found error");
                        }
                        else
                        {
                            throw new ApplicationException("Error code: " + httpResponse.StatusCode + " - " + text);
                        }
                    }
                }
            }
        }

  } // class

}

