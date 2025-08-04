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
                        Console.WriteLine("Id: " + items[m].Id);
                        Console.WriteLine("Subject: " + items[m].Subject);

                        Console.WriteLine("Plain body:");
                        Console.WriteLine(items[m].Body);
                        Console.WriteLine("-------------------------------------------------------");

                        Console.WriteLine("Html body:");
                        Console.WriteLine(items[m].BodyHtmlText);
                        Console.WriteLine("-------------------------------------------------------");

                        if (items[m].BodyRtf != null)
                        {
                            Console.WriteLine("Rtf body:");
                            Console.WriteLine(System.Text.Encoding.UTF8.GetString(items[m].BodyRtf));
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
