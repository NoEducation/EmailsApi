using Microsoft.Office.Interop.Outlook;
using OpenPop.Mime;
using OpenPop.Pop3;
using System;
using System.Collections.Generic;


// Microsoft.Office.Interop.Outlook wyglda że wymaga zainstlowej odpowiedniej wersji na maszynie.
namespace OutlookApi
{
    class Program
    {
        static void Main(string[] args)
        {
            Pop3Client pop3Client = new Pop3Client();//create an object for pop3client

            pop3Client.Connect("outlook.office365.com", 995, true);
            pop3Client.Authenticate("atrasik@office.wit.edu.pl", "*", AuthenticationMethod.UsernameAndPassword);

            int count = pop3Client.GetMessageCount(); //total count of email in MessageBox  ie our inbox
            var Emails = new List<POPClientEmail>(); //object for list POPClientEmail class which we already created. 

            int counter = 0;
            for (int i = count; i <= count ; i--)//going to read mails from last till total number of mails received
            {
                Message message = pop3Client.GetMessage(i);//assigning messagenumber to get detailed mail.//each mail having messagenumber

                POPClientEmail email = new POPClientEmail()
                {
                    MessageNumber = i,
                    Subject = message.Headers.Subject,
                    DateSent = message.Headers.DateSent,
                    From = message.Headers.From.Address,
                };
                MessagePart body = message.FindFirstHtmlVersion();
                if (body != null)
                {
                    email.Body = body.GetBodyAsText();
                }
                else
                {
                    body = message.FindFirstPlainTextVersion();
                    if (body != null)
                    {
                        email.Body = body.GetBodyAsText();
                    }
                }
                var attachments = message.FindAllAttachments();

                foreach (MessagePart attachment in attachments)
                {
                    email.Attachments.Add(new Attachment
                    {
                        FileName = attachment.FileName,
                        ContentType = attachment.ContentType.MediaType,
                        Content = attachment.Body
                    });
                }
                Emails.Add(email);
                counter++;

            }

            var emails = Emails;//You can filter mails by date from this list
        


        //try
        //{
        //    Application outlookApp = new Application();

        //    // Log in to Outlook
        //    NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
        //    outlookNamespace.Logon();

        //    // Get the Inbox folder
        //    MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

        //    // Retrieve the emails in the Inbox folder
        //    Items emails = inboxFolder.Items;

        //    foreach (MailItem email in emails)
        //    {
        //        try
        //        {
        //            Console.WriteLine("Subject: " + email.Subject);
        //            Console.WriteLine("Sender: " + email.SenderName);
        //            Console.WriteLine();
        //        }
        //        catch (System.Exception ex)
        //        {
        //            // Handle specific exceptions related to email processing
        //            Console.WriteLine("Error processing email: " + ex.Message);
        //        }
        //    }

        //    // Log out and release resources
        //    outlookNamespace.Logoff();
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookNamespace);
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);
        //}
        //catch (System.Exception ex)
        //{
        //    // Handle general exceptions related to Outlook connection or login
        //    Console.WriteLine("Error connecting to Outlook: " + ex.Message);
        //}
        //finally
        //{
        //    // Ensure resources are properly released even in case of exceptions
        //    GC.Collect();
        //    GC.WaitForPendingFinalizers();
        //    GC.Collect();
        //    GC.WaitForPendingFinalizers();
        //}

        //Console.WriteLine("Press any key to exit...");
        //Console.ReadKey();
    }
    }
}
