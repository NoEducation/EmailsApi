using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;

namespace ExchangeApi
{
    /// Can use Microsoft Graph https://developer.microsoft.com/en-us/graph/graph-explorer
    /// OR different approach https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            ExchangeService exchange = new ExchangeService(ExchangeVersion.Exchange2007_SP1);

            exchange.Credentials = new WebCredentials("atrasik", "*", "office.wit.edu.pl");
            exchange.AutodiscoverUrl("atrasik@office.wit.edu.pl");

            var result = exchange.
                FindItems(WellKnownFolderName.Inbox, new ItemView(100));

            var emails = new HashSet<ExchangeMail>();

            foreach (var item in result)
            {
                EmailMessage message = EmailMessage.Bind(exchange, item.Id);

                emails.Add(new ExchangeMail()
                {
                    Body = message.Body.Text,
                    From = message.From.Id.ToString(),
                    Subject = message.Subject.ToString(),
                });
            }

            foreach (var email in emails)
                Console.WriteLine($"{email.From},{email.Subject}.{email.Body}");
        }
    }
}
