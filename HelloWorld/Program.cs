﻿using System;
using System.Net;
using Microsoft.Exchange.WebServices.Data;

namespace HelloWorld
{
    class Program
    {
    static void Main(string[] args)
    {
    //   ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
      ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010);
    //   service.Credentials = new WebCredentials("user1@contoso.com", "password");
    //   service.Credentials = new WebCredentials("sven@htb.local", "Summer2020");
      service.Credentials = new WebCredentials("s.svensson", "Summer2020", "htb");
      service.TraceEnabled = true;
      service.TraceFlags = TraceFlags.All;

        // BypassSSLCert
        ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

    //   service.AutodiscoverUrl("sven@htb.local", RedirectionUrlValidationCallback);
      service.Url = new Uri("https://10.10.10.210/EWS/Exchange.asmx");
      EmailMessage email = new EmailMessage(service);
      email.ToRecipients.Add("sven@htb.local");
      email.Subject = "HelloWorld";
      email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API");
      email.Send();
    }
    private static bool RedirectionUrlValidationCallback(string redirectionUrl)
    {
      // The default for the validation callback is to reject the URL.
      bool result = false;
      Uri redirectionUri = new Uri(redirectionUrl);
      // Validate the contents of the redirection URL. In this simple validation
      // callback, the redirection URL is considered valid if it is using HTTPS
      // to encrypt the authentication credentials. 
      if (redirectionUri.Scheme == "https")
      {
        result = true;
      }
      return result;
    }
    }
}