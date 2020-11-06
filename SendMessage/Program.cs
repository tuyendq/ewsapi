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

            
                // Connect to Exchange Web Services as user1 at contoso.com.
                //    ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                // Create the binding
                ExchangeService service = new ExchangeService();
                if (service is null) {
                    System.Console.WriteLine("service variable is null.");
                } else {
                    System.Console.WriteLine("service variable created successfully!");
                }
                // Set the credentials for the on-pre server
                // service.Credentials = new WebCredentials("htb\\s.svensson", "Summer2020");
                service.Credentials = new WebCredentials("sven@htb.local", "Summer2020");

                // BypassCertificateError();
                // bool development = true;
                // ServicePointManager.ServerCertificateValidationCallback +=
                //     (sender, certificate, chain, errors) => {
                //         // local dev, just approve all certs
                //         if (development) return true;
                //         return errors == SslPolicyErrors.None ;
                //     };
                
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };  
                System.Console.WriteLine("1. Passed SSL Certificate Check!");
                // System.Net.ServicePointManager.ServerCertificateValidationCallback +=
                //     (s, cert, chain, sslPolicyErrors) => true;   
                // Using Autodiscovery
                // service.AutodiscoverUrl("sven@htb.local");
                // Set the URL manually
                service.Url = new Uri("https://10.10.10.210/EWS/Exchange.asmx");
                System.Console.WriteLine("2. Passed new Uri() Check!");


               // Create the e-mail message, set its properties, and send it to user2@contoso.com, saving a copy to the Sent Items folder. 
               EmailMessage message = new EmailMessage(service);
               System.Console.WriteLine("3. Passed new EmailMessage()");
               message.Subject = "Interesting";
               message.Body = "The proposition has been considered."; 
               message.ToRecipients.Add("sven@htb.local");
               System.Console.WriteLine("4. Passed message preparation");
               if (message is null) {
                   System.Console.WriteLine("message variable is null!");
               } else {
                   System.Console.WriteLine("Check message Subject: " + message.Subject);
                   System.Console.WriteLine("Before calling SendAndSaveCopy()");
                
                    // message.Send();
                    // System.Console.WriteLine(message.GetType());
                 
                    message.Save(); 

                    // message.SendAndSaveCopy();
                    Console.WriteLine("An email with the subject '" + message.Subject + "' has been sent to '" + message.ToRecipients[0] + "' and saved in the SendItems folder.");                   
               }
               
                //    System.Console.WriteLine("4.1 message variable: " + message);
               System.Console.WriteLine("5. Passed SendAndSaveCopy()");

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
            // ServicePointManager.ServerCertificateValidationCallback +=
            //     delegate(
            //         Object sender1,
            //         X509Certificate certificate,
            //         X509Chain chain,
            //         SslPolicyErrors sslPolicyErrors)
            //         {
            //             return true;
            //         };
            ServicePointManager.ServerCertificateValidationCallback +=
                (sender, certificate, chain, errors) => {
                    return true;
            };
        }
    }
}
