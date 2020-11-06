using System;
using Microsoft.Exchange.WebServices.Data;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates; 
using System.Net;

namespace SendMessage
{
    class Program
    {
        static void Main(string[] args)
            // Validate the server certificate.
            // For a certificate validation code example, see the Validating X509 Certificates topic in the Core Tasks section.
        {
            try 
            {
                BypassCertificateError();

                // Connect to Exchange Web Services as user1 at contoso.com.
                //    ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                // Create the binding
                ExchangeService service = new ExchangeService();
                // Set the credentials for the on-pre server
                service.Credentials = new WebCredentials("sven@htb.local", "Summer2020");
                // Using Autodiscovery
                // service.AutodiscoverUrl("sven@htb.local");
                // Set the URL manually
                service.Url = new Uri("https://10.10.10.210/EWS/Exchange.asmx");

               // Create the e-mail message, set its properties, and send it to user2@contoso.com, saving a copy to the Sent Items folder. 
               EmailMessage message = new EmailMessage(service);
               message.Subject = "Interesting";
               message.Body = "The proposition has been considered."; 
               message.ToRecipients.Add("sven@htb.local");
               message.SendAndSaveCopy();

               // Write confirmation message to console window.
               Console.WriteLine("Message sent!");
               Console.ReadLine();
            }
            catch (Exception ex)
            {
               Console.WriteLine("Error: " + ex.Message);
               Console.ReadLine();
            }
        }

        // Bypass CertificateError
        public static void BypassCertificateError()
        {
            ServicePointManager.ServerCertificateValidationCallback +=

                delegate(
                    Object sender1,
                    X509Certificate certificate,
                    X509Chain chain,
                    SslPolicyErrors sslPolicyErrors)
                    {
                        return true;
                    };
        }
    }
}
