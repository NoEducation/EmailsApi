using AE.Net.Mail;
using System;

namespace GmailAPI.SmptHelper
{
    public static class ImapClientHelper
    {
        public static void ReadFirstEmail(this ImapClient imapClient)
        {
            imapClient.SelectMailbox("INBOX");
            var email = imapClient.GetMessage(0);
            Console.WriteLine(email.Subject);
            imapClient.DeleteMessage(email); 
            Console.ReadLine();
        }

        public static ImapClient CreateClient() => new ImapClient("imap.gmail.com", "FootballNeighborhoodApp", "qpaeogsxkvxdzpjh", AuthMethods.Login, 993, true);
    }
}
