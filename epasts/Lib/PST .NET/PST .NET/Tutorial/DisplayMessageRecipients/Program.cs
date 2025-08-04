using System;
using System.Collections.Generic;
using Independentsoft.Pst;

namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            PstFile file = new PstFile("c:\\testfolder\\Outlook.pst");

            using (file)
            {
                Folder inbox = file.MailboxRoot.GetFolder("Inbox");

                if (inbox != null)
                {
                    IList<Item> items = inbox.GetItems();

                    for (int m = 0; m < items.Count; m++)
                    {
                        if (items[m] is Message)
                        {
                            Message message = (Message)items[m];

                            Console.WriteLine("Id: " + message.Id);
                            Console.WriteLine("Subject: " + message.Subject);

                            for (int r = 0; r < message.Recipients.Count; r++)
                            {
                                Recipient recipient = message.Recipients[r];

                                Console.WriteLine("Name: " + recipient.DisplayName);
                                Console.WriteLine("Email address: " + recipient.EmailAddress);
                                Console.WriteLine("Recipient type: " + recipient.RecipientType);
                            }

                            Console.WriteLine("-------------------------------------------------------");
                        }
                    }
                }
            }

            Console.WriteLine("Press ENTER to exit.");
            Console.Read();
        }
    }
}
